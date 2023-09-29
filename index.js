import * as core from '@actions/core';
import * as github from '@actions/github';
import fetch from 'node-fetch';

async function doFetch(url, options, expStatus, opDesc, respBody = true) {
  const response = await fetch(url, options);
  const jsonResponse = respBody? await response.json(): {};
  if (response.status != expStatus) {
    throw new Error("Received unexpected status code " + response.status + " on " + opDesc + " operation with response body: " + JSON.stringify(jsonResponse))
  }
  return jsonResponse;
}

async function createSession(token, sheetUrl, persistent) {
  const resp = await doFetch(sheetUrl + '/createSession',
    {method: 'POST',
     headers: {'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json'},
     body: JSON.stringify({persistChanges: persistent})},
    201, 'Create session');
  return resp.id;
}

async function closeSession(token, sheetUrl, sessionId) {
  await doFetch(sheetUrl + '/closeSession',
    {method: 'POST',
     headers: {'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json', 'workbook-session-id': sessionId}},
    204, 'Close session', false);
}

function toObj(headers, row) {
  const rowObj = {};
  for (var i = 0; i < headers.length; i++) {
      rowObj[headers[i]] = i < row.length? row[i]: '';
  }
  return rowObj;
}

async function getSheetData(token, sheetUrl, sheetName) {
  const rangeResp = await doFetch(sheetUrl + '/worksheets/' + sheetName + '/usedRange',
    {headers: {'Authorization': `Bearer ${token}`}}, 200, 'Fetch sheet data');
  const headers = rangeResp.values.shift();
  const data = rangeResp.values.map(row => toObj(headers, row));
  return [headers, data]
}

async function getUnprocessedIncomingRows(token, sheetUrl, statusCol) {
  const [headers, dataRows] = await getSheetData(token, sheetUrl, 'incoming');
  const result = [];
  for (var i = 0; i < dataRows.length; i++) {
    var row = dataRows[i];
    if (row[statusCol] === '') {
        result.push({rowIdx: i + 1, rowData: row});
    }
  }
  return [headers, result];
}

async function filterUnprocessedIncomingRows(token, sheetUrl, statusCol) {
  const table = 'intake_form';
  const sessionId = await createSession(token, sheetUrl, false);
  try {
    await doFetch(sheetUrl + '/tables/' + table + '/clearFilters',
      {method: 'POST', headers: {'Authorization': `Bearer ${token}`, 'Workbook-Session-Id': sessionId}}, 204, 'Clear table filters', false);
    await doFetch(sheetUrl + '/tables/' + table + '/columns/' + statusCol + '/filter/apply',
      {method: 'POST', headers: {'Authorization': `Bearer ${token}`, 'Workbook-Session-Id': sessionId},
       body: JSON.stringify({criteria: {filterOn: 'values', values: [''] } })}, 204, 'Apply status filter', false);
    const resp = await doFetch(sheetUrl + '/tables/' + table + '/range/visibleView/rows',
      {method: 'GET', headers: {'Authorization': `Bearer ${token}`, 'Workbook-Session-Id': sessionId}}, 200, 'Get filtered rows');
    const headers = resp.value.shift().values[0];
    const result = [];
    for (const row of resp.value) {
        result.push({rowAddresses: row.cellAddresses[0], rowData: toObj(headers, row.values[0])});
    }
    return [headers, result];
  } finally {
    await closeSession(token, sheetUrl, sessionId);
  }
}

async function getKV(token, sheetUrl, sheetName, keyCol, valueCol, allowEmptyVal, base) {
  const [headers, dataRows] = await getSheetData(token, sheetUrl, sheetName);
  const result = JSON.parse(JSON.stringify(base));
  for (const row of dataRows) {
      const val = row[valueCol];
      if ((val === '') && (! allowEmptyVal)) {
          continue;
      }
      result[row[keyCol]] = val;
  }
  return result;
}

function toRowNum(cellAddress) {
    var startIdx = -1;
    for (var i = 0; i < cellAddress.length; i++) {
        if (/^\d$/.test(cellAddress[i])) {
            startIdx = i;
            break;
        }
    };
    return parseInt(cellAddress.substring(startIdx));
}

function toContiguousBlocks(rowUpdates) {
    if (rowUpdates.length == 0) {
        return [];
    }
    const blocks = [];
    var rowNum = toRowNum(rowUpdates[0][0][0]);
    blocks.push([rowNum, rowNum, [rowUpdates[0]]]);
    for (var i = 1; i < rowUpdates.length; i++) {
        rowNum = toRowNum(rowUpdates[i][0][0]);
        if (rowNum == blocks[blocks.length - 1][1] + 1) {
            blocks[blocks.length - 1][1] = rowNum;
            blocks[blocks.length - 1][2].push(rowUpdates[i]);
        } else {
            blocks.push([rowNum, rowNum, [rowUpdates[i]]]);
        }
    }
    return blocks;
}

async function updateTableRow(token, sessionId, sheetUrl, table, numCols, tableRowIdx, updateIdx, updateVals) {
    const vals = new Array(numCols).fill(null);
    for (var i = 0; i < updateIdx.length; i++) {
        vals[updateIdx[i]] = updateVals[i];
    }
    await doFetch(`${sheetUrl}/tables/${table}/rows/$/ItemAt(index=${tableRowIdx})`,
        {method: 'PATCH', body: JSON.stringify({values: [vals]}),
         headers: {'Authorization': `Bearer ${token}`, 'Workbook-Session-Id': sessionId}}, 200, 'Update table row');
}

async function updateRowsViaBlockRange(token, sessionId, sheetUrl, sheet, updateIdx, block) {
    const values = [];
    for (const [rowAddr, updateVals] of block) {
        const vals = new Array(rowAddr.length).fill(null);
        for (var i = 0; i < updateIdx.length; i++) {
            vals[updateIdx[i]] = updateVals[i];
        }
        values.push(vals);
    }
    const topLeft = block[0][0][0];
    const lastRowAddr = block[block.length - 1][0];
    const bottomRight = lastRowAddr[lastRowAddr.length - 1];
    console.log(`Making batch update call for rows ${topLeft}:${bottomRight}`);
    await doFetch(`${sheetUrl}/worksheets/${sheet}/range(address='${topLeft}:${bottomRight}')`,
        {method: 'PATCH', body: JSON.stringify({values}),
         headers: {'Authorization': `Bearer ${token}`, 'Workbook-Session-Id': sessionId, 'Content-Type': 'application/json'}},
        200, 'Update range');
}

async function getSFToken(instanceUrl, clientId, clientSecret) {
  const url = `${instanceUrl}/services/oauth2/token?grant_type=client_credentials&client_id=${clientId}&client_secret=${clientSecret}`
  const resp = await doFetch(url, {method: 'POST'}, 200, "SalesForce token generation")
  return resp.access_token;
}

async function createSFRecord(token, instanceUrl, version, recordType, row, fieldMapping) {
  const url = `${instanceUrl}/services/data/${version}/sobjects/${recordType}`
  const body = {}
  for (const [sheetField, sfField] of Object.entries(fieldMapping)) {
    body[sfField] = row[sheetField];
  }
  const options = {method: 'POST', body: JSON.stringify(body),
                   headers: {'Content-Type': 'application/json', 'Authorization': `Bearer ${token}`}}
  const resp = await doFetch(url, options, 201, "SalesForce record creation");
  return resp;
}


try {
  const payload = github.context.payload.client_payload;
  console.log('Event payload: ', JSON.stringify(payload));
  const sourceLocation = payload.sourceLocation;

  const authToken = core.getInput('ms-graph-auth-token');
  const sfInstanceUrl = core.getInput('sf-instance-url');
  const sfClientId = core.getInput('sf-client-id');
  const sfClientSecret = core.getInput('sf-client-secret');
  const defaultConfig = JSON.parse(core.getInput('default-settings'));

  async function doWork() {
    if (! sourceLocation.startsWith('onedrive:/')) {
        throw new Error('Unsupported source location ' + sourceLocation);
    }
    const workbookUrl = `https://graph.microsoft.com/v1.0/${sourceLocation.substring('onedrive:/'.length)}/workbook`
    console.log("Base config: ", JSON.stringify(defaultConfig));
    const metadata = await getKV(authToken, workbookUrl, 'metadata', 'Key', 'Value', true, defaultConfig);
    console.log("Final config: ", JSON.stringify(metadata))
    const statusCol = metadata['IncomingStatusColumn'];
    const msgCol = metadata['IncomingErrorMsgColumn'];
    const sfType = metadata['SFObjectType'];
    const sfVersion = metadata['SFObjectVersion'];
    const sfMappingSheet = metadata['SFMappingSheet'];
    const formFieldNameCol = metadata['FormFieldNameColumn'];
    const sfFieldNameCol = metadata['SFFieldNameColumn'];
    const sfMapping = await getKV(authToken, workbookUrl, sfMappingSheet, formFieldNameCol, sfFieldNameCol, false, {});
    console.log("Mapping: ", JSON.stringify(sfMapping));

    const [headers, rows] = await filterUnprocessedIncomingRows(authToken, workbookUrl, statusCol);
    const statusIdx = headers.indexOf(statusCol);
    const msgIdx = headers.indexOf(msgCol);
    const token = await getSFToken(sfInstanceUrl, sfClientId, sfClientSecret)
    const data = []
    const result = []
    const updates = []
    for (const row of rows) {
        try {
            const rowRes = await createSFRecord(token, sfInstanceUrl, sfVersion, sfType, row.rowData, sfMapping);
            result.push({rowRange: `${row.rowAddresses[0]}:${row.rowAddresses[row.rowAddresses.length - 1]}`, status: 'SUCCESS', response: rowRes});
            updates.push([row.rowAddresses, ['Y', 'SUCCESS']]);
        } catch (err) {
            result.push({rowRange: `${row.rowAddresses[0]}:${row.rowAddresses[row.rowAddresses.length - 1]}`, status: 'FAILURE', errorMessage: err.message});
            updates.push([row.rowAddresses, ['N', `FAILURE: ${err.message}`]]);
        }
    }
    const sessionId = await createSession(authToken, workbookUrl, true);
    try {
        const blocks = toContiguousBlocks(updates);
        for (const block of blocks) {
            await updateRowsViaBlockRange(authToken, sessionId, workbookUrl, 'incoming', [statusIdx, msgIdx], block[2]);
        }
    } finally {
      await closeSession(authToken, workbookUrl, sessionId);
    }
    return result;
  }
  doWork()
    .then(result => {
      console.log('Invocation result: ', JSON.stringify({result}));
      core.setOutput("processing-result", JSON.stringify({result}));
    })
    .catch(err => {
      core.setFailed(err.message);
      throw err;
    })
} catch (error) {
  core.setFailed(error.message);
  throw error;
}

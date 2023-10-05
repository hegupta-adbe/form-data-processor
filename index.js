import * as core from '@actions/core';
import * as github from '@actions/github';
import fetch from 'node-fetch';
import _sodium from 'libsodium-wrappers';

async function doFetchRaw(url, options, expStatus, opDesc, respBody = true) {
  const response = await fetch(url, options);
  var jsonResponse = {};
  try {
    jsonResponse = await response.json();
  } catch (err) {
    if (respBody) {
      console.log(`WARN: Received unexpected non-JSON response on operation ${opDesc}`);
    }
  }
  if (response.status != expStatus) {
    throw new Error("Received unexpected status code " + response.status + " on " + opDesc + " operation with response body: " + JSON.stringify(jsonResponse))
  }
  return jsonResponse;
}

async function encryptForGithub(secret, key) {
  await _sodium.ready;
  const sodium = _sodium;
  let binkey = sodium.from_base64(key, sodium.base64_variants.ORIGINAL)
  let binsec = sodium.from_string(secret)
  let encBytes = sodium.crypto_box_seal(binsec, binkey)
  let output = sodium.to_base64(encBytes, sodium.base64_variants.ORIGINAL)
  return output
}

async function updateGithubSecret(tokenHolder, secretName, secretVal) {
  const encSecret = await encryptForGithub(secretVal, tokenHolder.repoKey);
  const url = `https://api.github.com/repos/${tokenHolder.repoOwnerAndName}/actions/secrets/${secretName}`;
  const headers = {'Accept': 'application/vnd.github+json', 'Authorization': `Bearer ${tokenHolder.repoToken}`, 'X-GitHub-Api-Version': '2022-11-28'};
  const body = JSON.stringify({encrypted_value: encSecret, key_id: tokenHolder.repoKeyID});
  await doFetchRaw(url, {method: 'PUT', headers, body}, 204, `Update Github secret ${secretName}`, false);
}

async function refreshTokens(tokenHolder) {
  const formData = {};
  formData.client_id = tokenHolder.graphClientId;
  formData.scope = tokenHolder.graphScope;
  formData.refresh_token = tokenHolder.graphRefreshToken;
  formData.grant_type = 'refresh_token';
  formData.client_secret = tokenHolder.graphClientSecret;
  var formBody = [];
  for (var property in formData) {
    var encodedKey = encodeURIComponent(property);
    var encodedValue = encodeURIComponent(formData[property]);
    formBody.push(encodedKey + "=" + encodedValue);
  }
  formBody = formBody.join("&");
  const response = await doFetchRaw(`https://login.microsoftonline.com/${tokenHolder.graphTenant}/oauth2/v2.0/token`,
    {method: 'POST', headers: {'Content-Type': 'application/x-www-form-urlencoded'}, body: formBody},
    200, 'Refresh Graph tokens');
  tokenHolder.graphAccessToken = response.access_token;
  await updateGithubSecret(tokenHolder, tokenHolder.repoGraphAccessTokenSecret, tokenHolder.graphAccessToken);
  tokenHolder.graphRefreshToken = response.refresh_token;
  await updateGithubSecret(tokenHolder, tokenHolder.repoGraphRefreshTokenSecret, tokenHolder.graphRefreshToken);
}

async function doFetchWithRefresh(url, tokenHolder, options, expStatus, opDesc, respBody = true) {
  if (! ('headers' in options)) {
    options.headers = {};
  }
  options.headers.Authorization = `Bearer ${tokenHolder.graphAccessToken}`;
  try {
    return await doFetchRaw(url, options, expStatus, opDesc, respBody);
  } catch (err) {
    // TODO do this better!
    if (err.message.includes('Received unexpected status code 401 on ') && err.message.includes('"InvalidAuthenticationToken"')) {
      console.log("Detected invalid access token, trying to refresh");
      await refreshTokens(tokenHolder);
      options.headers.Authorization = `Bearer ${tokenHolder.graphAccessToken}`;
      return await doFetchRaw(url, options, expStatus, opDesc, respBody);
    } else {
      throw err;
    }
  }
}

async function createSession(tokenHolder, sheetUrl, persistent) {
  const resp = await doFetchWithRefresh(sheetUrl + '/createSession', tokenHolder,
    {method: 'POST',
     headers: {'Content-Type': 'application/json'},
     body: JSON.stringify({persistChanges: persistent})},
    201, 'Create session');
  return resp.id;
}

async function closeSession(tokenHolder, sheetUrl, sessionId) {
  await doFetchWithRefresh(sheetUrl + '/closeSession', tokenHolder,
    {method: 'POST',
     headers: {'Content-Type': 'application/json', 'workbook-session-id': sessionId}},
    204, 'Close session', false);
}

function toObj(headers, row) {
  const rowObj = {};
  for (var i = 0; i < headers.length; i++) {
      rowObj[headers[i]] = i < row.length? row[i]: '';
  }
  return rowObj;
}

async function getSheetData(tokenHolder, sheetUrl, sheetName) {
  const rangeResp = await doFetchWithRefresh(sheetUrl + '/worksheets/' + sheetName + '/usedRange', tokenHolder,
    {}, 200, 'Fetch sheet data');
  const headers = rangeResp.values.shift();
  const data = rangeResp.values.map(row => toObj(headers, row));
  return [headers, data]
}

async function getUnprocessedIncomingRows(tokenHolder, sheetUrl, statusCol) {
  const [headers, dataRows] = await getSheetData(tokenHolder, sheetUrl, 'incoming');
  const result = [];
  for (var i = 0; i < dataRows.length; i++) {
    var row = dataRows[i];
    if (row[statusCol] === '') {
        result.push({rowIdx: i + 1, rowData: row});
    }
  }
  return [headers, result];
}

async function filterUnprocessedIncomingRows(tokenHolder, sheetUrl, statusCol) {
  const table = 'intake_form';
  const sessionId = await createSession(tokenHolder, sheetUrl, false);
  try {
    await doFetchWithRefresh(sheetUrl + '/tables/' + table + '/clearFilters', tokenHolder,
      {method: 'POST', headers: {'Workbook-Session-Id': sessionId}}, 204, 'Clear table filters', false);
    await doFetchWithRefresh(sheetUrl + '/tables/' + table + '/columns/' + statusCol + '/filter/apply', tokenHolder,
      {method: 'POST', headers: {'Workbook-Session-Id': sessionId},
       body: JSON.stringify({criteria: {filterOn: 'values', values: [''] } })}, 204, 'Apply status filter', false);
    const resp = await doFetchWithRefresh(sheetUrl + '/tables/' + table + '/range/visibleView/rows', tokenHolder,
      {method: 'GET', headers: {'Workbook-Session-Id': sessionId}}, 200, 'Get filtered rows');
    const headers = resp.value.shift().values[0];
    const result = [];
    for (const row of resp.value) {
        result.push({rowAddresses: row.cellAddresses[0], rowData: toObj(headers, row.values[0])});
    }
    return [headers, result];
  } finally {
    await closeSession(tokenHolder, sheetUrl, sessionId);
  }
}

async function getKV(tokenHolder, sheetUrl, sheetName, keyCol, valueCol, allowEmptyVal, base) {
  const [headers, dataRows] = await getSheetData(tokenHolder, sheetUrl, sheetName);
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

async function updateTableRow(tokenHolder, sessionId, sheetUrl, table, numCols, tableRowIdx, updateIdx, updateVals) {
    const vals = new Array(numCols).fill(null);
    for (var i = 0; i < updateIdx.length; i++) {
        vals[updateIdx[i]] = updateVals[i];
    }
    await doFetchWithRefresh(`${sheetUrl}/tables/${table}/rows/$/ItemAt(index=${tableRowIdx})`, tokenHolder,
        {method: 'PATCH', body: JSON.stringify({values: [vals]}),
         headers: {'Workbook-Session-Id': sessionId}}, 200, 'Update table row');
}

async function updateRowsViaBlockRange(tokenHolder, sessionId, sheetUrl, sheet, updateIdx, block) {
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
    await doFetchWithRefresh(`${sheetUrl}/worksheets/${sheet}/range(address='${topLeft}:${bottomRight}')`, tokenHolder,
        {method: 'PATCH', body: JSON.stringify({values}),
         headers: {'Workbook-Session-Id': sessionId, 'Content-Type': 'application/json'}},
        200, 'Update range');
}

async function getSFToken(instanceUrl, clientId, clientSecret) {
  const url = `${instanceUrl}/services/oauth2/token?grant_type=client_credentials&client_id=${clientId}&client_secret=${clientSecret}`
  const resp = await doFetchRaw(url, {method: 'POST'}, 200, "SalesForce token generation")
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
  const resp = await doFetchRaw(url, options, 201, "SalesForce record creation");
  return resp;
}


try {
  const payload = github.context.payload.client_payload;
  console.log('Event payload: ', JSON.stringify(payload));
  const sourceLocation = payload.sourceLocation;

  const repoOwnerAndName = core.getInput('repo-owner-name');
  const repoKeyID = core.getInput('repo-key-id');
  const repoKey = core.getInput('repo-key');
  const repoToken = core.getInput('repo-token');
  const repoGraphAccessTokenSecret = core.getInput('repo-graph-access-token-secret');
  const repoGraphRefreshTokenSecret = core.getInput('repo-graph-refresh-token-secret');

  const graphTenant = core.getInput('ms-graph-tenant');
  const graphAccessToken = core.getInput('ms-graph-auth-token');
  const graphRefreshToken = core.getInput('ms-graph-refresh-token');
  const graphClientId = core.getInput('ms-graph-client-id');
  const graphClientSecret = core.getInput('ms-graph-client-secret');
  const graphScope = core.getInput('ms-graph-scope');

  const sfInstanceUrl = core.getInput('sf-instance-url');
  const sfClientId = core.getInput('sf-client-id');
  const sfClientSecret = core.getInput('sf-client-secret');

  const defaultConfig = JSON.parse(core.getInput('default-settings'));

  async function doWork() {
    if (! sourceLocation.startsWith('onedrive:/')) {
        throw new Error('Unsupported source location ' + sourceLocation);
    }
    const tokenholder = {graphTenant, graphAccessToken, graphRefreshToken, graphClientId, graphClientSecret, graphScope,
        repoOwnerAndName, repoKeyID, repoKey, repoToken, repoGraphAccessTokenSecret, repoGraphRefreshTokenSecret};
    const workbookUrl = `https://graph.microsoft.com/v1.0/${sourceLocation.substring('onedrive:/'.length)}/workbook`
    console.log("Base config: ", JSON.stringify(defaultConfig));
    const metadata = await getKV(tokenholder, workbookUrl, 'metadata', 'Key', 'Value', true, defaultConfig);
    console.log("Final config: ", JSON.stringify(metadata))
    const statusCol = metadata['IncomingStatusColumn'];
    const msgCol = metadata['IncomingErrorMsgColumn'];
    const sfType = metadata['SFObjectType'];
    const sfVersion = metadata['SFObjectVersion'];
    const sfMappingSheet = metadata['SFMappingSheet'];
    const formFieldNameCol = metadata['FormFieldNameColumn'];
    const sfFieldNameCol = metadata['SFFieldNameColumn'];
    const sfMapping = await getKV(tokenholder, workbookUrl, sfMappingSheet, formFieldNameCol, sfFieldNameCol, false, {});
    console.log("Mapping: ", JSON.stringify(sfMapping));

    const [headers, rows] = await filterUnprocessedIncomingRows(tokenholder, workbookUrl, statusCol);
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
    const sessionId = await createSession(tokenholder, workbookUrl, true);
    try {
        const blocks = toContiguousBlocks(updates);
        for (const block of blocks) {
            await updateRowsViaBlockRange(tokenholder, sessionId, workbookUrl, 'incoming', [statusIdx, msgIdx], block[2]);
        }
    } finally {
      await closeSession(tokenholder, workbookUrl, sessionId);
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

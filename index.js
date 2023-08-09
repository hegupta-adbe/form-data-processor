import * as core from '@actions/core';
import * as github from '@actions/github';
import fetch from 'node-fetch';
import {google} from 'googleapis';

const sheets = google.sheets('v4');

function getBatchUpdateRequest(auth, spreadsheetId, data) {
    return {
        spreadsheetId,
        resource: {
          valueInputOption: 'RAW',
          data
        },
        auth
    };
}

function appendIncomingCellUpdate(data, rowIdx, colIdx, val) {
    const cellR1C1 = "R" + (rowIdx + 1) + "C" + (colIdx + 1);
    data.push({range: "incoming!" + cellR1C1 + ":" + cellR1C1, values: [[val]]});
}

async function getOrCreateCol(headers, auth, spreadsheetId, colName) {
    var colIdx = headers.indexOf(colName)
    if (colIdx === -1) {
        headers.push(colName);
        colIdx = headers.length - 1;
        const data = []
        appendIncomingCellUpdate(data, 0, colIdx, colName);
        const request = getBatchUpdateRequest(auth, spreadsheetId, data);
        await sheets.spreadsheets.values.batchUpdate(request);
    }
    return colIdx;
}

async function setupIncoming(auth, spreadsheetId, statusCol, msgCol) {
  const getResult = await sheets.spreadsheets.values.get({
      auth,
      spreadsheetId,
      range: 'incoming!1:1'
    });
  const headers = getResult.data.values[0];
  const statusColIdx = await getOrCreateCol(headers, auth, spreadsheetId, statusCol);
  const msgColIdx = await getOrCreateCol(headers, auth, spreadsheetId, msgCol);
  return {headers, statusColIdx, msgColIdx}
}

async function getUnprocessedIncomingRows(auth, spreadsheetId, statusCol) {
  const getResult = await sheets.spreadsheets.values.get({
      auth,
      spreadsheetId,
      range: 'incoming'
    });
  const allRowsWithHeader = getResult.data.values;
  const headers = allRowsWithHeader.shift();
  const dataRows = allRowsWithHeader.map(row => {
    const rowObj = {};
    for (var i = 0; i < headers.length; i++) {
        rowObj[headers[i]] = i < row.length? row[i]: '';
    }
    return rowObj;
  });
  const result = [];
  for (var i = 0; i < dataRows.length; i++) {
    var row = dataRows[i];
    if (row[statusCol] === '') {
        result.push({rowIdx: i + 1, rowData: row});
    }
  }
  return result;
}

async function doFetch(url, options, expStatus, opDesc) {
  const response = await fetch(url, options);
  const jsonResponse = await response.json();
  if (response.status != expStatus) {
    throw new Error("Received unexpected status code " + response.status + " on " + opDesc + " operation with response body: " + JSON.stringify(jsonResponse))
  }
  return jsonResponse;
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
  const saEmail = core.getInput('google-sa-email');
  const saPK = core.getInput('google-sa-pk');
  const spreadsheetId = core.getInput('google-sheet-id');
  const statusCol = core.getInput('sheet-status-column');
  const msgCol = core.getInput('sheet-error-column');

  const sfInstanceUrl = core.getInput('sf-instance-url');
  const sfClientId = core.getInput('sf-client-id');
  const sfClientSecret = core.getInput('sf-client-secret');
  const sfVersion = core.getInput('sf-version');
  const sfType = core.getInput('sf-type');

  const sfMapping = JSON.parse(core.getInput('form-to-sf-mapping'));

  const auth = new google.auth.JWT(
       saEmail,
       null,
       saPK,
       ['https://www.googleapis.com/auth/spreadsheets']);
  async function doWork() {
    const headerResp = await setupIncoming(auth, spreadsheetId, statusCol, msgCol);
    const rows = await getUnprocessedIncomingRows(auth, spreadsheetId, statusCol);
    const token = await getSFToken(sfInstanceUrl, sfClientId, sfClientSecret)
    const data = []
    const result = []
    for (const row of rows) {
        try {
            const rowRes = await createSFRecord(token, sfInstanceUrl, sfVersion, sfType, row.rowData, sfMapping);
            result.push({rowNumber: row.rowIdx + 1, status: 'SUCCESS', response: rowRes});
            appendIncomingCellUpdate(data, row.rowIdx, headerResp.statusColIdx, 'SUCCESS');
            appendIncomingCellUpdate(data, row.rowIdx, headerResp.msgColIdx, '')
        } catch (err) {
            result.push({rowNumber: row.rowIdx + 1, status: 'FAILURE', errorMessage: err.message});
            appendIncomingCellUpdate(data, row.rowIdx, headerResp.statusColIdx, 'FAILURE')
            appendIncomingCellUpdate(data, row.rowIdx, headerResp.msgColIdx, err.message)
        }
    }
    const request = getBatchUpdateRequest(auth, spreadsheetId, data);
    await sheets.spreadsheets.values.batchUpdate(request);
    return result;
  }
  doWork()
    .then(result => {
      console.log('Invocation result: ', JSON.stringify({result}));
      core.setOutput("processing-result", JSON.stringify({result}));
    })
    .catch(err => {
      core.setFailed(err.message);
    })
} catch (error) {
  core.setFailed(error.message);
}

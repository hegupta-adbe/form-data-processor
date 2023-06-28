import * as core from '@actions/core';
import * as github from '@actions/github';
import fetch from 'node-fetch';
import {google} from 'googleapis';

const sheets = google.sheets('v4');

async function doFetch(url, options) {
  const response = await fetch(url, options);
  const jsonResponse = await response.json();
  return jsonResponse;
}

async function getSheetData(saEmail, saPK, spreadsheetId, range) {
  const auth = new google.auth.JWT(
       saEmail,
       null,
       saPK,
       ['https://www.googleapis.com/auth/spreadsheets.readonly']);
  const getResult = await sheets.spreadsheets.values.batchGet({
      auth,
      spreadsheetId,
      ranges: ['incoming!1:1', range]
    });
  const headers = getResult.data.valueRanges[0].values[0];
  const result = getResult.data.valueRanges[1].values.map(row => {
    const rowObj = {};
    for (var i = 0; i < headers.length; i++) {
        rowObj[headers[i]] = row[i];
    }
    return rowObj;
  });
  return result;
}

async function getSFToken(instanceUrl, clientId, clientSecret) {
  const url = `${instanceUrl}/services/oauth2/token?grant_type=client_credentials&client_id=${clientId}&client_secret=${clientSecret}`
  const resp = await doFetch(url, {method: 'POST'})
  return resp.access_token;
}

async function createSFRecord(token, instanceUrl, version, recordType, row, fieldMapping) {
  const url = `${instanceUrl}/services/data/${version}/sobjects/${recordType}`
  const body = {}
  for (const [sheetField, sfField] of Object.entries(fieldMapping)) {
    body[sfField] = row[sheetField];
  }
  const options = {body: JSON.stringify(body),
                   headers: {'Content-Type': 'application/json', 'Authorization': `Bearer ${token}`}}
  const resp = await doFetch(url, options);
  return resp;
}

try {
  const saEmail = core.getInput('google-sa-email');
  const saPK = core.getInput('google-sa-pk');
  const sfInstanceUrl = core.getInput('sf-instance-url');
  const sfClientId = core.getInput('sf-client-id');
  const sfClientSecret = core.getInput('sf-client-secret');
  const payload = github.context.payload.client_payload;
  console.log('Event payload: ', payload)
  // getSheetData(saEmail, saPK, '1aRD-ulu0YyNa9aNLg7zjVVYC5cRy9KPtBHlLtOXcgwc', 'incoming!12:13')
  // getSFToken(sfInstanceUrl, sfClientId, sfClientSecret)
  createSFRecord("00D010000004bAI!AQIAQIoa58V0tn3MnND_2rJ4OsXwM.V44VPEB8lAUzzaFnTWvBEhJJXBhKQ5MlrxctCR4yUeNhRdiVD.vYgW.F0W62QDVt6e",
                 sfInstanceUrl, "v57.0", "vega__c",
                 {fname: "Foo Bar", email: "foobar@company.com"}, {fname: "Name", email: "Email__c"})
    .then(result => {
      console.log('Invocation result: ', result);
      core.setOutput("processing-result", JSON.stringify({result}));
    })
    .catch(err => {
      core.setFailed(err.message);
    })
} catch (error) {
  core.setFailed(error.message);
}

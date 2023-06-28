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

try {
  const saEmail = core.getInput('google-sa-email');
  const saPK = core.getInput('google-sa-pk');
  const sfInstanceUrl = core.getInput('sf-instance-url');
  const sfClientId = core.getInput('sf-client-id');
  const sfClientSecret = core.getInput('sf-client-secret');
  const payload = github.context.payload.client_payload;
  console.log('Event payload: ', payload)
  // getSheetData(saEmail, saPK, '1aRD-ulu0YyNa9aNLg7zjVVYC5cRy9KPtBHlLtOXcgwc', 'incoming!12:13')
  getSFToken(sfInstanceUrl, sfClientId, sfClientSecret)
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

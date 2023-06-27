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
  const getResult = await sheets.spreadsheets.values.get({
      auth,
      spreadsheetId,
      range
    });
  return getResult.data.values;
}

try {
  const saEmail = core.getInput('google-sa-email');
  const saPK = core.getInput('google-sa-pk');
  const payload = github.context.payload.client_payload;
  console.log('Event payload: ', payload)
  getSheetData(saEmail, saPK, '1aRD-ulu0YyNa9aNLg7zjVVYC5cRy9KPtBHlLtOXcgwc', 'incoming!12:13')
    .then(jsonResponse => {
      console.log('Fetch result: ', jsonResponse);
      core.setOutput("processing-result", JSON.stringify(jsonResponse));
    })
    .catch(err => {
      core.setFailed(err.message);
    })
} catch (error) {
  core.setFailed(error.message);
}

const core = require('@actions/core');
const github = require('@actions/github');
import fetch from 'node-fetch';

async function doFetch(url, options) {
  const response = await fetch(url, options);
  const jsonResponse = await response.json();
  return jsonResponse;
}

try {
  const clientId = core.getInput('azure-client-id');
  const clientSecret = core.getInput('azure-client-secret');
  const resource = core.getInput('azure-resource');
  const payload = github.context.payload.client_payload
  const resp = {clientId, clientSecret, resource, payload}
  const options = {
    method: 'POST',
    body: JSON.stringify(resp),
    headers: { 'Content-Type': 'application/json' }
  }
  doFetch('https://echo.zuplo.io', options)
    .then(jsonResponse => {
      console.log(jsonResponse);
      core.setOutput("processing-result", jsonResponse);
    })
    .catch(err => {
      core.setFailed(err.message);
    })
} catch (error) {
  core.setFailed(error.message);
}

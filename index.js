import * as core from '@actions/core';
import * as github from '@actions/github';
import fetch from 'node-fetch';
import _sodium from 'libsodium-wrappers';
import { OneDrive, OneDriveAuth } from "@adobe/helix-onedrive-support";

async function doFetchRaw(url, options, expStatuses, opDesc, respBody = true) {
  const response = await fetch(url, options);
  var jsonResponse = {};
  try {
    jsonResponse = await response.json();
  } catch (err) {
    if (respBody) {
      console.log(`WARN: Received unexpected non-JSON response on operation ${opDesc}`);
    }
  }
  if (! expStatuses.includes(response.status)) {
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

async function updateGithubSecret(repoKeyID, repoKey, repoOwnerAndName, repoToken, secretName, secretVal) {
  const encSecret = await encryptForGithub(secretVal, repoKey);
  const url = `https://api.github.com/repos/${repoOwnerAndName}/actions/secrets/${secretName}`;
  const headers = {'Accept': 'application/vnd.github+json', 'Authorization': `Bearer ${repoToken}`, 'X-GitHub-Api-Version': '2022-11-28'};
  const body = JSON.stringify({encrypted_value: encSecret, key_id: repoKeyID});
  await doFetchRaw(url, {method: 'PUT', headers, body}, [201, 204], `Update Github secret ${secretName}`, false);
}

async function getSheetData(workbook, sheetName) {
  return await workbook.worksheet(sheetName).usedRange().getRowsAsObjects();
}

async function filterUnprocessedIncomingRows(workbook, statusCol, maxRows) {
  const table = workbook.table('intake_form');
  await workbook.createSession(false);
  try {
    await table.clearFilters();
    await table.applyFilter(statusCol, {filterOn: 'values', values: [''] });
    return await table.getVisibleRowsAsObjectsWithAddresses(maxRows);
  } finally {
    await workbook.closeSession();
  }
}

async function getKV(workbook, sheetName, keyCol, valueCol, allowEmptyVal, base) {
  const dataRows = await getSheetData(workbook, sheetName);
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

async function updateRowsViaBlockRange(workbook, sheet, updateIdx, block) {
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
    const range = workbook.worksheet(sheet).range(`${topLeft}:${bottomRight}`);
    await range.update(JSON.stringify({values}));
}

async function getSFToken(instanceUrl, clientId, clientSecret) {
  const url = `${instanceUrl}/services/oauth2/token?grant_type=client_credentials&client_id=${clientId}&client_secret=${clientSecret}`
  const resp = await doFetchRaw(url, {method: 'POST'}, [200], "SalesForce token generation")
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
  const resp = await doFetchRaw(url, options, [201], "SalesForce record creation");
  return resp;
}

const payload = github.context.payload.client_payload;
console.log('Event payload: ', JSON.stringify(payload));
const sourceLocation = payload.sourceLocation;

const repoOwnerAndName = core.getInput('repo-owner-name');
const repoKeyID = core.getInput('repo-key-id');
const repoKey = core.getInput('repo-key');
const repoToken = core.getInput('repo-token');
const repoMsalCacheDumpSecret = core.getInput('repo-msal-cache-dump-secret');

const graphTenant = core.getInput('ms-graph-tenant');
const graphRefreshToken = core.getInput('ms-graph-refresh-token');
const graphClientId = core.getInput('ms-graph-client-id');
const graphClientSecret = core.getInput('ms-graph-client-secret');
const graphScopes = core.getInput('ms-graph-scope').split(' ');
const maxRows = parseInt(core.getInput('max-rows'));
var msalCacheDump = core.getInput('msal-cache-dump');

const sfInstanceUrl = core.getInput('sf-instance-url');
const sfClientId = core.getInput('sf-client-id');
const sfClientSecret = core.getInput('sf-client-secret');

const defaultConfig = JSON.parse(core.getInput('default-settings'));

class GithubCachePlugin {

    async beforeCacheAccess(cacheContext) {
      console.log("LOAD CALLBACK");
      if (msalCacheDump) {
        cacheContext.tokenCache.deserialize(msalCacheDump);
      }
    }

    async afterCacheAccess(cacheContext) {
      console.log("SAVE CALLBACK");
      if (cacheContext.cacheHasChanged) {
        console.log("WRITING MODIFIED CACHE");
        msalCacheDump = cacheContext.tokenCache.serialize();
        await updateGithubSecret(repoKeyID, repoKey, repoOwnerAndName, repoToken, repoMsalCacheDumpSecret, msalCacheDump);
      }
    }

    async getPluginMetadata() {
      return {};
    }
}

try {
  async function doWork() {
    if (! sourceLocation.startsWith('onedrive:/')) {
        throw new Error('Unsupported source location ' + sourceLocation);
    }

    const cachePlugin = new GithubCachePlugin();
    const auth = new OneDriveAuth({
      clientId: graphClientId,
      clientSecret: graphClientSecret,
      tenant: graphTenant,
      cachePlugin,
      scopes: graphScopes
    });
    if (! msalCacheDump) {
      await auth.app.acquireTokenByRefreshToken({refreshToken: graphRefreshToken, scopes: graphScopes, forceCache: true});
    }
    const client = new OneDrive({
      auth,
    });
    const workbook = client.getWorkbookFromPath(sourceLocation.substring('onedrive:'.length) + '/workbook');

    console.log("Base config: ", JSON.stringify(defaultConfig));
    const metadata = await getKV(workbook, 'metadata', 'Key', 'Value', true, defaultConfig);
    console.log("Final config: ", JSON.stringify(metadata));
    const statusCol = metadata['IncomingStatusColumn'];
    const msgCol = metadata['IncomingErrorMsgColumn'];
    const sfType = metadata['SFObjectType'];
    const sfVersion = metadata['SFObjectVersion'];
    const sfMappingSheet = metadata['SFMappingSheet'];
    const formFieldNameCol = metadata['FormFieldNameColumn'];
    const sfFieldNameCol = metadata['SFFieldNameColumn'];
    const sfMapping = await getKV(workbook, sfMappingSheet, formFieldNameCol, sfFieldNameCol, false, {});
    console.log("SF Field Mapping: ", JSON.stringify(sfMapping));

    const headers = await workbook.table('intake_form').getHeaderNames();
    const rows = await filterUnprocessedIncomingRows(workbook, statusCol, maxRows);
    const statusIdx = headers.indexOf(statusCol);
    const msgIdx = headers.indexOf(msgCol);
    const token = await getSFToken(sfInstanceUrl, sfClientId, sfClientSecret)
    const data = []
    const result = []
    const updates = []
    for (const row of rows) {
        try {
            const rowRes = await createSFRecord(token, sfInstanceUrl, sfVersion, sfType, row.data, sfMapping);
            result.push({rowRange: `${row.cellAddresses[0]}:${row.cellAddresses[row.cellAddresses.length - 1]}`, status: 'SUCCESS', response: rowRes});
            updates.push([row.cellAddresses, ['Y', 'SUCCESS']]);
        } catch (err) {
            result.push({rowRange: `${row.cellAddresses[0]}:${row.cellAddresses[row.cellAddresses.length - 1]}`, status: 'FAILURE', errorMessage: err.message});
            updates.push([row.cellAddresses, ['N', `FAILURE: ${err.message}`]]);
        }
    }
    await workbook.createSession(true);
    try {
        const blocks = toContiguousBlocks(updates);
        for (const block of blocks) {
            await updateRowsViaBlockRange(workbook, 'incoming', [statusIdx, msgIdx], block[2]);
        }
    } finally {
      await workbook.closeSession();
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

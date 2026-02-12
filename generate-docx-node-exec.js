// ============================================================
// NODE: "Generate DOCX Report" — n8n Code Node (exec workaround)
// ============================================================
// Mode: Run Once for All Items
// Place AFTER: "Format Slack + Report Data"
//
// No NODE_FUNCTION_ALLOW_EXTERNAL needed.
// No template literals — copy/paste safe.
// ============================================================

var childProcess = require('child_process');
var fs = require('fs');

var items = $input.all();
var analysis = items[0].json.analysis;
var payload = items[0].json.payload || {};
var slackMessage = items[0].json.slackMessage;

// Build the input JSON for the standalone script
var scriptInput = JSON.stringify({
  analysis: analysis,
  metadata: {
    portalId: (payload.metadata && payload.metadata.portalId) ? payload.metadata.portalId : 'N/A',
    tier: (payload.metadata && payload.metadata.tier) ? payload.metadata.tier : 'HubSpot',
    dataSource: (payload.metadata && payload.metadata.dataSource) ? payload.metadata.dataSource : 'API + CSV'
  }
});

// Path to the script on your Render filesystem
// Update this path if you placed it elsewhere
var SCRIPT_PATH = '/home/node/campaign-audit-report.js';

try {
  // Escape single quotes in the JSON for shell safety
  var safeInput = scriptInput.replace(/'/g, "'\\''");

  // Execute the script, piping JSON via stdin
  var result = childProcess.execSync(
    "echo '" + safeInput + "' | node \"" + SCRIPT_PATH + "\"",
    {
      timeout: 30000,
      maxBuffer: 10 * 1024 * 1024,
      encoding: 'utf8'
    }
  );

  // Parse the script output (JSON with filepath, filename, size)
  var output = JSON.parse(result.trim());

  // Read the generated DOCX file
  var fileBuffer = fs.readFileSync(output.filepath);

  // Return json + binary
  return [{
    json: {
      slackMessage: slackMessage,
      analysis: analysis,
      payload: payload,
      reportGenerated: true,
      reportFilename: output.filename,
      reportSize: output.size
    },
    binary: {
      report: await this.helpers.prepareBinaryData(
        fileBuffer,
        output.filename,
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      )
    }
  }];

} catch (err) {
  // If DOCX generation fails, still let the Slack text message through
  return [{
    json: {
      slackMessage: slackMessage,
      analysis: analysis,
      payload: payload,
      reportGenerated: false,
      reportError: err.message,
      reportFilename: null
    }
  }];
}

/**
 * Khanin Diagram — Google Sheets Add-on
 * (c) 2026 Daniil Khanin and Khanin Solutions S.L.
 * License: BSL 1.1
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Khanin Diagram')
    .addItem('Open Diagram', 'showDiagram')
    .addItem('Create Template', 'createTemplate')
    .addToUi();
}

function onInstall(e) {
  onOpen();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function showDiagram() {
  var html = HtmlService.createTemplateFromFile('Dialog')
    .evaluate()
    .setWidth(900)
    .setHeight(700)
    .setTitle('Khanin Diagram');
  SpreadsheetApp.getUi().showModalDialog(html, 'Khanin Diagram');
}

/**
 * Creates template sheets with example data
 */
function createTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Settings sheet ---
  var settingsSheet = ss.getSheetByName('KD_Settings');
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('KD_Settings');
  }
  settingsSheet.clear();
  settingsSheet.getRange('A1:B1').setValues([['Parameter', 'Value']]).setFontWeight('bold');
  settingsSheet.getRange('A2:B8').setValues([
    ['Container Label', 'MAU'],
    ['Container Value', 2150000],
    ['Container Delta', -3.2],
    ['Center Label', 'Revenue'],
    ['Center Value', 720000],
    ['Center Delta', 1.2],
    ['Center Format', 'money']
  ]);
  settingsSheet.autoResizeColumns(1, 2);

  // --- Zones sheet ---
  var zonesSheet = ss.getSheetByName('KD_Zones');
  if (!zonesSheet) {
    zonesSheet = ss.insertSheet('KD_Zones');
  }
  zonesSheet.clear();
  zonesSheet.getRange('A1:E1').setValues([['Zone ID', 'Label', 'Arc Start', 'Arc End', 'Color']]).setFontWeight('bold');
  zonesSheet.getRange('A2:E3').setValues([
    ['supply', 'SUPPLY', 180, 360, '#A89B7E'],
    ['demand', 'DEMAND', 0, 180, '#3A4047']
  ]);
  zonesSheet.autoResizeColumns(1, 5);

  // --- Nodes sheet ---
  var nodesSheet = ss.getSheetByName('KD_Nodes');
  if (!nodesSheet) {
    nodesSheet = ss.insertSheet('KD_Nodes');
  }
  nodesSheet.clear();
  nodesSheet.getRange('A1:F1').setValues([['Node ID', 'Label', 'Value', 'Delta %', 'Zone', 'Group']]).setFontWeight('bold');
  nodesSheet.getRange('A2:F10').setValues([
    ['sellers',      'Sellers',      98400,   -5.6,  'supply', 'output'],
    ['freeSellers',  'Free Sellers', 62700,   -7.3,  'supply', 'supply'],
    ['adsFree',      'Ads Free',     215300,  -2.8,  'supply', 'supply'],
    ['leadsFree',    'Leads Free',   612500,  -16.4, 'supply', 'merge'],
    ['customers',    'Customers',    35700,   4.5,   'demand', 'demand'],
    ['adsPaid',      'Ads Paid',     44800,   -8.9,  'demand', 'demand'],
    ['leadsPaid',    'Leads Paid',   488200,  -13.7, 'demand', 'merge'],
    ['leadsTotal',   'Leads Total',  1100700, -15.3, 'demand', 'output'],
    ['transactions', 'Transactions', 184600,  2.1,   'demand', 'revenue']
  ]);
  nodesSheet.autoResizeColumns(1, 6);

  // --- Flows sheet ---
  var flowsSheet = ss.getSheetByName('KD_Flows');
  if (!flowsSheet) {
    flowsSheet = ss.insertSheet('KD_Flows');
  }
  flowsSheet.clear();
  flowsSheet.getRange('A1:D1').setValues([['From', 'To', 'Value', 'Group']]).setFontWeight('bold');
  flowsSheet.getRange('A2:D13').setValues([
    ['container',   'sellers',      98400,   'supply'],
    ['sellers',     'freeSellers',  62700,   'supply'],
    ['sellers',     'customers',    35700,   'demand'],
    ['freeSellers', 'adsFree',      215300,  'supply'],
    ['adsFree',     'leadsFree',    612500,  'supply'],
    ['customers',   'adsPaid',      44800,   'demand'],
    ['adsPaid',     'leadsPaid',    488200,  'demand'],
    ['container',   'leadsTotal',   1100700, 'direct'],
    ['leadsTotal',  'leadsFree',    612500,  'merge'],
    ['leadsTotal',  'leadsPaid',    488200,  'merge'],
    ['adsPaid',     'transactions', 184600,  'revenue'],
    ['transactions','center',       720000,  'revenue']
  ]);
  flowsSheet.autoResizeColumns(1, 4);

  SpreadsheetApp.getUi().alert(
    'Template created!\n\n' +
    'Sheets added: KD_Settings, KD_Zones, KD_Nodes, KD_Flows\n\n' +
    'Edit the data, then open: Khanin Diagram → Open Diagram'
  );
}

/**
 * Reads diagram config from sheets. Called from client-side JS.
 */
function getDiagramData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Settings
  var settingsSheet = ss.getSheetByName('KD_Settings');
  if (!settingsSheet) {
    return { error: 'Sheet "KD_Settings" not found. Use "Create Template" first.' };
  }

  var settings = {};
  var sData = settingsSheet.getDataRange().getValues();
  for (var i = 1; i < sData.length; i++) {
    settings[sData[i][0]] = sData[i][1];
  }

  // Zones
  var zonesSheet = ss.getSheetByName('KD_Zones');
  var zones = [];
  if (zonesSheet) {
    var zData = zonesSheet.getDataRange().getValues();
    for (var j = 1; j < zData.length; j++) {
      if (!zData[j][0]) continue;
      zones.push({
        id: String(zData[j][0]),
        label: String(zData[j][1]),
        arc: [Number(zData[j][2]), Number(zData[j][3])],
        color: String(zData[j][4] || '#888888')
      });
    }
  }

  // Nodes
  var nodesSheet = ss.getSheetByName('KD_Nodes');
  var nodes = [];
  if (nodesSheet) {
    var nData = nodesSheet.getDataRange().getValues();
    for (var k = 1; k < nData.length; k++) {
      if (!nData[k][0]) continue;
      nodes.push({
        id: String(nData[k][0]),
        label: String(nData[k][1]),
        value: Number(nData[k][2]),
        delta: Number(nData[k][3]),
        zone: String(nData[k][4]),
        group: String(nData[k][5]) || undefined
      });
    }
  }

  // Flows
  var flowsSheet = ss.getSheetByName('KD_Flows');
  var flows = [];
  if (flowsSheet) {
    var fData = flowsSheet.getDataRange().getValues();
    for (var m = 1; m < fData.length; m++) {
      if (!fData[m][0]) continue;
      flows.push({
        from: String(fData[m][0]),
        to: String(fData[m][1]),
        value: Number(fData[m][2]),
        group: String(fData[m][3]) || undefined
      });
    }
  }

  return {
    container: {
      label: settings['Container Label'] || 'MAU',
      value: Number(settings['Container Value']) || 0,
      delta: Number(settings['Container Delta']) || 0
    },
    center: {
      label: settings['Center Label'] || 'Revenue',
      value: Number(settings['Center Value']) || 0,
      delta: Number(settings['Center Delta']) || 0,
      format: settings['Center Format'] || undefined
    },
    zones: zones,
    nodes: nodes,
    flows: flows,
    animation: { enabled: false }
  };
}

/**
 * Inserts SVG as an image into the active sheet.
 * Google Sheets doesn't support SVG directly, so we convert to PNG via Charts.
 */
function insertDiagramImage(dataUrl) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var blob = Utilities.newBlob(
    Utilities.base64Decode(dataUrl.split(',')[1]),
    'image/png',
    'khanin-diagram.png'
  );
  sheet.insertImage(blob, 1, 1);
}

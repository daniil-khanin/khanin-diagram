/**
 * Khanin Diagram — Google Sheets Add-on
 * (c) 2026 Daniil Khanin and Khanin Solutions S.L.
 * License: BSL 1.1
 *
 * Simple Sankey-like input: one sheet, columns From / To / Value / Delta %
 * Everything else (zones, groups, layout) is auto-detected.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Khanin Diagram')
    .addItem('Open Diagram', 'showDiagram')
    .addItem('Create Template', 'createTemplate')
    .addToUi();
}

function onInstall(e) { onOpen(); }

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

// ─── Template ────────────────────────────────────────────────────────────────

function createTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Khanin Diagram');
  if (!sheet) sheet = ss.insertSheet('Khanin Diagram');
  sheet.clear();

  sheet.getRange('A1:D1')
    .setValues([['From', 'To', 'Value', 'Delta %']])
    .setFontWeight('bold')
    .setBackground('#f3f3f3');

  sheet.getRange('A2:D14').setValues([
    ['MAU',          'MAU',           2150000, -3.2],
    ['MAU',          'Sellers',         98400, -5.6],
    ['Sellers',      'Free Sellers',    62700, -7.3],
    ['Sellers',      'Customers',       35700,  4.5],
    ['Free Sellers', 'Ads Free',       215300, -2.8],
    ['Ads Free',     'Leads Free',     612500, -16.4],
    ['Customers',    'Ads Paid',        44800, -8.9],
    ['Ads Paid',     'Leads Paid',     488200, -13.7],
    ['MAU',          'Leads Total',   1100700, -15.3],
    ['Leads Total',  'Leads Free',     612500,  ''],
    ['Leads Total',  'Leads Paid',     488200,  ''],
    ['Ads Paid',     'Transactions',   184600,  2.1],
    ['Transactions', 'Revenue',        720000,  1.2]
  ]);

  // Add note to self-reference rows
  sheet.getRange('A2').setNote('Self-reference row (From = To) sets the container value and delta');

  sheet.autoResizeColumns(1, 4);

  SpreadsheetApp.getUi().alert(
    'Template created!\n\n' +
    'Sheet "Khanin Diagram" with example data.\n\n' +
    'How it works:\n' +
    '• Each row is a flow: From → To with a Value\n' +
    '• Row where From = To sets that node\'s own value & delta\n' +
    '  (e.g. MAU → MAU = 2.15M is the container total)\n' +
    '• Delta % is optional\n' +
    '• Container (root) and Center (target) are auto-detected\n\n' +
    'Edit the data, then: Khanin Diagram → Open Diagram'
  );
}

// ─── Data reading & auto-detection ───────────────────────────────────────────

function getDiagramData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Khanin Diagram');
  if (!sheet) {
    return { error: 'Sheet "Khanin Diagram" not found. Use "Create Template" first.' };
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return { error: 'No data found. Add rows with From, To, Value columns.' };
  }

  // ── 1. Parse rows ──────────────────────────────────────────────────────────
  var metadata = {};   // self-ref rows: name → { value, delta }
  var rawFlows = [];
  var fromSet = {};
  var toSet = {};

  for (var i = 1; i < data.length; i++) {
    var from = String(data[i][0]).trim();
    var to   = String(data[i][1]).trim();
    var val  = Number(data[i][2]) || 0;
    var dRaw = data[i][3];
    var delta = (dRaw !== '' && dRaw !== null && dRaw !== undefined) ? Number(dRaw) : null;
    if (isNaN(delta)) delta = null;

    if (!from || !to) continue;

    if (from === to) {
      metadata[from] = { value: val, delta: delta || 0 };
    } else {
      rawFlows.push({ from: from, to: to, value: val, delta: delta });
      fromSet[from] = true;
      toSet[to] = true;
    }
  }

  if (rawFlows.length === 0) {
    return { error: 'No flow data found. Add rows with From → To → Value.' };
  }

  // ── 2. Auto-detect container & center ──────────────────────────────────────
  var containerName = null;
  for (var fn in fromSet) {
    if (!toSet[fn]) { containerName = fn; break; }
  }
  if (!containerName) containerName = String(data[1][0]).trim();

  var centerName = null;
  for (var tn in toSet) {
    if (!fromSet[tn]) { centerName = tn; break; }
  }

  // Container value & delta
  var cm = metadata[containerName] || {};
  var containerValue = cm.value || 0;
  var containerDelta = cm.delta || 0;
  if (!containerValue) {
    rawFlows.forEach(function(f) { if (f.from === containerName) containerValue += f.value; });
  }

  // Center value & delta
  var centerValue = 0, centerDelta = 0;
  if (centerName) {
    var cenm = metadata[centerName] || {};
    centerValue = cenm.value || 0;
    centerDelta = cenm.delta || 0;
    if (!centerValue) {
      rawFlows.forEach(function(f) { if (f.to === centerName) centerValue += f.value; });
    }
  }

  // Auto-detect money format for center
  var moneyWords = ['revenue','income','profit','gmv','sales','earning','arpu','money'];
  var centerFormat = null;
  if (centerName) {
    var cl = centerName.toLowerCase();
    for (var mw = 0; mw < moneyWords.length; mw++) {
      if (cl.indexOf(moneyWords[mw]) >= 0) { centerFormat = 'money'; break; }
    }
  }

  // ── 3. Build adjacency (preserving table order) ────────────────────────────
  var adj = {};
  rawFlows.forEach(function(f) {
    if (!adj[f.from]) adj[f.from] = [];
    adj[f.from].push(f.to);
  });

  // ── 4. BFS depth from container ────────────────────────────────────────────
  var depth = {};
  depth[containerName] = 0;
  var queue = [containerName];
  var bfsVisited = {};
  bfsVisited[containerName] = true;
  while (queue.length > 0) {
    var cur = queue.shift();
    var nb = adj[cur] || [];
    for (var ni = 0; ni < nb.length; ni++) {
      if (!bfsVisited[nb[ni]]) {
        bfsVisited[nb[ni]] = true;
        depth[nb[ni]] = (depth[cur] || 0) + 1;
        queue.push(nb[ni]);
      }
    }
  }

  // ── 5. Auto-detect zones via DFS coloring ──────────────────────────────────
  //   At each branching node the FIRST child (table order) → supply,
  //   second child → demand. This lets the user control sides by row order.
  var nodeZone = {};

  function dfsColor(node, zone) {
    if (nodeZone[node] !== undefined) return;
    if (node === centerName || node === containerName) return;
    nodeZone[node] = zone;
    var children = adj[node] || [];
    var first = true;
    for (var ci = 0; ci < children.length; ci++) {
      if (nodeZone[children[ci]] !== undefined || children[ci] === centerName) continue;
      if (first) {
        dfsColor(children[ci], zone);
        first = false;
      } else {
        dfsColor(children[ci], zone === 'supply' ? 'demand' : zone);
      }
    }
  }

  var containerChildren = adj[containerName] || [];
  var firstChild = true;
  for (var cci = 0; cci < containerChildren.length; cci++) {
    if (containerChildren[cci] === centerName) continue;
    dfsColor(containerChildren[cci], firstChild ? 'supply' : 'demand');
    firstChild = false;
  }

  // ── 6. Collect all node names (except container & center) ──────────────────
  var allNodeNames = {};
  rawFlows.forEach(function(f) {
    if (f.from !== containerName) allNodeNames[f.from] = true;
    if (f.to !== centerName)     allNodeNames[f.to]   = true;
  });
  delete allNodeNames[containerName];

  // Ensure every node has a zone
  for (var an in allNodeNames) {
    if (!nodeZone[an]) nodeZone[an] = 'demand';
  }

  // ── 7. Detect groups ──────────────────────────────────────────────────────
  var incomingFrom = {};
  rawFlows.forEach(function(f) {
    if (!incomingFrom[f.to]) incomingFrom[f.to] = [];
    incomingFrom[f.to].push(f.from);
  });

  var containerDirect = {};
  rawFlows.forEach(function(f) {
    if (f.from === containerName) containerDirect[f.to] = true;
  });

  var nodeGroup = {};
  for (var gn in allNodeNames) {
    // Revenue: connects to center
    var connectsToCenter = false;
    (adj[gn] || []).forEach(function(t) { if (t === centerName) connectsToCenter = true; });
    if (connectsToCenter) { nodeGroup[gn] = 'revenue'; continue; }

    // Merge: receives 2+ incoming flows
    var inc = incomingFrom[gn] || [];
    if (inc.length >= 2) { nodeGroup[gn] = 'merge'; continue; }

    // Output: direct child of container
    if (containerDirect[gn]) { nodeGroup[gn] = 'output'; continue; }

    // Default: same as zone
    nodeGroup[gn] = nodeZone[gn] || 'supply';
  }

  // ── 8. Node values & deltas ────────────────────────────────────────────────
  var nodeValues = {};
  var nodeDeltas = {};
  rawFlows.forEach(function(f) {
    if (f.to === centerName) return;
    if (nodeValues[f.to] === undefined) nodeValues[f.to] = f.value;
    if (nodeDeltas[f.to] === undefined && f.delta !== null && f.delta !== 0) {
      nodeDeltas[f.to] = f.delta;
    }
  });

  // ── 9. Build nodes array ──────────────────────────────────────────────────
  var nodes = [];
  for (var nodeId in allNodeNames) {
    nodes.push({
      id:    nodeId,
      label: nodeId,
      value: nodeValues[nodeId] || 0,
      delta: nodeDeltas[nodeId] || 0,
      zone:  nodeZone[nodeId]  || 'supply',
      group: nodeGroup[nodeId] || 'supply'
    });
  }

  // ── 10. Build flows array ─────────────────────────────────────────────────
  var flows = [];
  rawFlows.forEach(function(f) {
    var fFrom = (f.from === containerName) ? 'container' : f.from;
    var fTo   = (f.to === centerName)      ? 'center'    : f.to;

    // Flow group detection
    var fg = 'supply';
    if (fTo === 'center' || nodeGroup[f.from] === 'revenue') {
      fg = 'revenue';
    } else if (fFrom === 'container' && nodeZone[f.to] === 'demand') {
      fg = 'direct';
    } else if (nodeGroup[f.from] === 'output' && nodeZone[f.from] === 'demand' && nodeGroup[f.to] === 'merge') {
      fg = 'merge';
    } else {
      fg = nodeZone[f.to] || 'supply';
    }

    flows.push({ from: fFrom, to: fTo, value: f.value, group: fg });
  });

  // ── 11. Generate layout ───────────────────────────────────────────────────
  var layout = generateLayout(nodes, depth, containerName, centerName, adj, nodeGroup);

  return {
    container: {
      label: containerName,
      value: containerValue,
      delta: containerDelta
    },
    center: centerName ? {
      label:  centerName,
      value:  centerValue,
      delta:  centerDelta,
      format: centerFormat
    } : null,
    zones: [
      { id: 'supply', label: 'SUPPLY', arc: [180, 360], color: '#A89B7E' },
      { id: 'demand', label: 'DEMAND', arc: [0, 180],   color: '#3A4047' }
    ],
    nodes:  nodes,
    flows:  flows,
    layout: layout,
    animation: { enabled: false }
  };
}

// ─── Auto-layout ─────────────────────────────────────────────────────────────

function generateLayout(nodes, depth, containerName, centerName, adj, nodeGroup) {
  var maxDepth = 1;
  for (var k in depth) {
    if (k !== containerName && k !== centerName && depth[k] > maxDepth) maxDepth = depth[k];
  }

  // Separate by zone, sort by depth
  var supply = [], demand = [];
  nodes.forEach(function(n) {
    (n.zone === 'supply' ? supply : demand).push(n);
  });
  supply.sort(function(a, b) { return (depth[a.id] || 1) - (depth[b.id] || 1); });
  demand.sort(function(a, b) { return (depth[a.id] || 1) - (depth[b.id] || 1); });

  var layout = {};

  function placeZone(nodeList, ySign) {
    // Group by depth
    var byDepth = {};
    nodeList.forEach(function(n) {
      var d = depth[n.id] || 1;
      if (!byDepth[d]) byDepth[d] = [];
      byDepth[d].push(n);
    });

    var depths = Object.keys(byDepth).map(Number).sort(function(a, b) { return a - b; });

    for (var di = 0; di < depths.length; di++) {
      var d = depths[di];
      var nodesAtD = byDepth[d];
      var t = (d - 1) / Math.max(maxDepth - 1, 1);

      var baseX = -0.50 + t * 0.90;
      var baseY = ySign * (0.12 + t * 0.38);

      for (var ni = 0; ni < nodesAtD.length; ni++) {
        var xSpread = 0;
        if (nodesAtD.length > 1) {
          xSpread = (ni - (nodesAtD.length - 1) / 2) * 0.20;
        }
        layout[nodesAtD[ni].id] = {
          x: baseX + xSpread,
          y: baseY + ni * 0.06 * ySign,
          labelAbove: ySign < 0
        };
      }
    }
  }

  placeZone(supply, -1);
  placeZone(demand,  1);

  // Adjust special groups
  nodes.forEach(function(n) {
    var pos = layout[n.id];
    if (!pos) return;

    if (n.group === 'output') {
      // Output nodes that feed merge nodes → far right
      var children = adj[n.id] || [];
      var feedsMerge = children.some(function(c) { return nodeGroup[c] === 'merge'; });
      if (feedsMerge) {
        pos.x = 0.68;
        pos.y = 0;
        pos.labelAbove = true;
      }
      // Otherwise keep depth-based position (entry output node)
    } else if (n.group === 'revenue') {
      pos.x = Math.max(pos.x, 0.32);
      pos.y *= 0.35;
      pos.labelAbove = pos.y <= 0;
    } else if (n.group === 'merge') {
      pos.y *= 0.88;
    }
  });

  // Force relaxation — push apart overlapping nodes
  var ids = Object.keys(layout);
  for (var iter = 0; iter < 6; iter++) {
    for (var a = 0; a < ids.length; a++) {
      for (var b = a + 1; b < ids.length; b++) {
        var p1 = layout[ids[a]], p2 = layout[ids[b]];
        var dx = p2.x - p1.x, dy = p2.y - p1.y;
        var dd = Math.sqrt(dx * dx + dy * dy);
        if (dd < 0.22 && dd > 0.001) {
          var push = (0.22 - dd) * 0.3;
          var nx = dx / dd, ny = dy / dd;
          p1.x -= nx * push; p1.y -= ny * push;
          p2.x += nx * push; p2.y += ny * push;
        }
      }
    }
    for (var c = 0; c < ids.length; c++) {
      var p = layout[ids[c]];
      var r = Math.sqrt(p.x * p.x + p.y * p.y);
      if (r > 0.72) { p.x *= 0.72 / r; p.y *= 0.72 / r; }
    }
  }

  return layout;
}

// ─── Image insertion ─────────────────────────────────────────────────────────

function insertDiagramImage(dataUrl) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var blob = Utilities.newBlob(
    Utilities.base64Decode(dataUrl.split(',')[1]),
    'image/png',
    'khanin-diagram.png'
  );
  sheet.insertImage(blob, 1, 1);
}

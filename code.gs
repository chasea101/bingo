// Bingo - Google Apps Script backend
//
// SETUP:
//   1. Create a new Google Sheet
//   2. Extensions > Apps Script, paste this file, save
//   3. Deploy > New deployment > Web app
//      Execute as: Me, Who has access: Anyone
//   4. Copy the Web App URL into API_URL in index.html
//
// REDEPLOY after edits: Deploy > Manage deployments > edit > New version

var LETTERS = 'ABCDEFGHJKLMNPQRSTUVWXYZ';
var MAX_ITEMS = 100;
var MAX_ITEM_LENGTH = 60;

function doGet(e) {
  var p = (e && e.parameter) || {};
  var result;

  switch (p.action || '') {
    case 'createGame': result = handleCreate(p.name); break;
    case 'join':       result = handleJoin(p.code, p.name); break;
    case 'getState':   result = handleGetState(p.code); break;
    case 'addItem':    result = handleAddItem(p.code, p.name, p.text); break;
    case 'removeItem': result = handleRemoveItem(p.code, p.name, p.id); break;
    case 'lock':       result = handleLock(p.code, p.name); break;
    case 'reopen':     result = handleReopen(p.code, p.name); break;
    case 'toggle':     result = handleToggle(p.code, p.name, p.id); break;
    default:           result = { error: 'Unknown action' };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- Storage ---

function readState(code) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(code);
  if (!sheet) return null;
  var raw = sheet.getRange('A1').getValue();
  return raw ? JSON.parse(raw) : null;
}

function writeState(code, state) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(code)
    .getRange('A1').setValue(JSON.stringify(state));
}

function generateCode() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  for (var attempt = 0; attempt < 100; attempt++) {
    var code = '';
    for (var i = 0; i < 4; i++) {
      code += LETTERS.charAt(Math.floor(Math.random() * LETTERS.length));
    }
    if (!ss.getSheetByName(code)) return code;
  }
  return null;
}

// Runs fn(state) under a lock. fn returns an error object or null to save.
function mutate(code, fn) {
  if (!code) return { error: 'Missing code' };
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var state = readState(code);
    if (!state) return { error: 'Game not found' };
    var err = fn(state);
    if (err) return err;
    writeState(code, state);
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

// --- Handlers ---

function handleCreate(name) {
  name = (name || '').trim();
  if (!name) return { error: 'Missing name' };
  var code = generateCode();
  if (!code) return { error: 'Could not generate room code' };

  SpreadsheetApp.getActiveSpreadsheet().insertSheet(code);
  writeState(code, {
    phase: 'building',
    host: name,
    players: [name],
    items: [],
    nextId: 1
  });
  return { success: true, code: code };
}

function handleJoin(code, name) {
  name = (name || '').trim();
  if (!name) return { error: 'Missing name' };
  return mutate(code, function(state) {
    if (state.players.indexOf(name) === -1) state.players.push(name);
  });
}

function handleGetState(code) {
  if (!code) return { error: 'Missing code' };
  var state = readState(code);
  return state || { error: 'Game not found' };
}

function handleAddItem(code, name, text) {
  text = (text || '').trim().replace(/\s+/g, ' ');
  if (!text) return { error: 'Empty item' };
  if (text.length > MAX_ITEM_LENGTH) return { error: 'Max ' + MAX_ITEM_LENGTH + ' characters' };
  return mutate(code, function(state) {
    if (state.players.indexOf(name) === -1) return { error: 'Not in this game' };
    if (state.phase !== 'building') return { error: 'Board is locked' };
    if (state.items.length >= MAX_ITEMS) return { error: 'Board is full' };
    var lower = text.toLowerCase();
    for (var i = 0; i < state.items.length; i++) {
      if (state.items[i].text.toLowerCase() === lower) return { error: 'Already on the board' };
    }
    state.items.push({ id: state.nextId++, text: text, by: name, marked: false, markedBy: null });
  });
}

function handleRemoveItem(code, name, id) {
  id = parseInt(id, 10);
  return mutate(code, function(state) {
    if (state.phase !== 'building') return { error: 'Board is locked' };
    for (var i = 0; i < state.items.length; i++) {
      if (state.items[i].id === id) {
        if (state.items[i].by !== name && state.host !== name) return { error: 'Not your item' };
        state.items.splice(i, 1);
        return null;
      }
    }
    return { error: 'Item not found' };
  });
}

function handleLock(code, name) {
  return mutate(code, function(state) {
    if (state.host !== name) return { error: 'Only the host can do this' };
    if (state.items.length < 4) return { error: 'Need at least 4 items' };
    // shuffle
    var a = state.items;
    for (var i = a.length - 1; i > 0; i--) {
      var j = Math.floor(Math.random() * (i + 1));
      var t = a[i]; a[i] = a[j]; a[j] = t;
    }
    state.phase = 'playing';
  });
}

function handleReopen(code, name) {
  return mutate(code, function(state) {
    if (state.host !== name) return { error: 'Only the host can do this' };
    state.phase = 'building';
    state.items.forEach(function(it) { it.marked = false; it.markedBy = null; });
  });
}

function handleToggle(code, name, id) {
  id = parseInt(id, 10);
  return mutate(code, function(state) {
    if (state.players.indexOf(name) === -1) return { error: 'Not in this game' };
    if (state.phase !== 'playing') return { error: 'Board is not locked yet' };
    for (var i = 0; i < state.items.length; i++) {
      var it = state.items[i];
      if (it.id === id) {
        it.marked = !it.marked;
        it.markedBy = it.marked ? name : null;
        return null;
      }
    }
    return { error: 'Item not found' };
  });
}

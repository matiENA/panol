/**
 * PANOL CLOUD - MASTER SYSTEM v6.3
 * Enhanced Analytics with Brand/Product data
 */

const CONFIG = {
  SHEETS: {
    ITEMS: 'DB_ITEMS', 
    TRANSACTIONS: 'DB_TRANSACTIONS',
    STAFF: 'DB_STAFF',
    OT: 'DB_OT_LIST',
    MESSAGING: 'DB_MESSAGING'
  },
  STAFF_COLS: { ID: 0, NAME: 1, ROLE: 2, SCORE: 3, AVG_TIME: 4, LAST_ACTIVE: 5, ZONE: 6 },
  OT_COLS: { UNIT_ID: 0, OT_NUMBER: 1, SEMI: 2, OT_NUMBER_ALT: 3, PRODUCT: 4, BRAND: 5, BRAND_SEMI: 6 },
  MSG_COLS: { MSG_ID: 0, TIMESTAMP: 1, SENDER: 2, TARGET_OP_ID: 3, UNIT_CONTEXT: 4, TYPE: 5, BODY: 6, STATUS: 7, REPLY: 8, SCORE: 9, TAG: 10 }
};

function getSS() { return SpreadsheetApp.getActiveSpreadsheet(); }
function _getSheet(name) { var ss = getSS(); var sheet = ss.getSheetByName(name); if (!sheet) sheet = ss.insertSheet(name); return sheet; }

function doGet(e) {
  var route = e && e.parameter && e.parameter.v ? e.parameter.v : 'home';
  var title = 'App Mecánicos';
  var file = 'Index'; 
  if (route === 'dashboard') { file = 'Index-dashboard'; title = 'Command Center'; }
  if (route === 'panol') { file = 'Index-panol'; title = 'Monitor Pañol'; }
  if (route === 'inv') { file = 'Index-inv'; title = 'Control de Stock'; }
  return HtmlService.createHtmlOutputFromFile(file).setTitle(title).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// === MESSAGING ===
function sendDispatchCard(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const sheet = _getSheet(CONFIG.SHEETS.MESSAGING);
    const msgId = "MSG-" + Math.floor(Math.random() * 1000000);
    if (!payload.targetOpId || !payload.body) return { success: false, error: "Missing required fields" };
    sheet.appendRow([msgId, new Date(), "Dashboard", String(payload.targetOpId).trim(), payload.unitId || "GENERAL", payload.type || "REQUEST", payload.body, "UNREAD", "", 0, ""]);
    return { success: true, msgId: msgId };
  } catch (e) { return { success: false, error: e.message }; } 
  finally { lock.releaseLock(); }
}

function pollInbox(opId) {
  const sheet = _getSheet(CONFIG.SHEETS.MESSAGING);
  const data = sheet.getDataRange().getValues();
  const messages = [];
  const target = String(opId).trim();
  const startRow = Math.max(1, data.length - 100);
  for (let i = startRow; i < data.length; i++) {
    const row = data[i];
    if (String(row[3]).trim() === target && String(row[7]).trim() === "UNREAD" && String(row[5]).trim() !== "RATING_LOG") {
      messages.push({ id: row[0], sender: row[2], unit: row[4], type: row[5], body: row[6] });
    }
  }
  return messages;
}

function resolveFeedback(msgId, responseText) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const sheet = _getSheet(CONFIG.SHEETS.MESSAGING);
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(msgId)) {
        sheet.getRange(i + 1, 8).setValue("RESOLVED");
        sheet.getRange(i + 1, 9).setValue(responseText || "Sin respuesta");
        return { success: true };
      }
    }
    return { success: false, error: "Message not found" };
  } catch (e) { return { success: false, error: e.message }; } 
  finally { lock.releaseLock(); }
}

// === RATING ===
function rateInteraction(refId, scoreDelta, tag) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const msgSheet = _getSheet(CONFIG.SHEETS.MESSAGING);
    const txSheet = _getSheet(CONFIG.SHEETS.TRANSACTIONS);
    let targetOpId = null;
    
    const msgData = msgSheet.getDataRange().getValues();
    const shadowId = "LOG-" + refId;
    
    // Check if rating exists
    for(let i = msgData.length - 1; i >= 1; i--) {
      if(String(msgData[i][0]) === shadowId) {
        msgSheet.getRange(i+1, 10).setValue(scoreDelta);
        msgSheet.getRange(i+1, 11).setValue(tag);
        targetOpId = msgData[i][3];
        if(targetOpId) updateMechanicScore(targetOpId);
        return { success: true };
      }
    }
    
    // Create new rating from transaction
    const txData = txSheet.getDataRange().getValues();
    for(let i = txData.length - 1; i >= 1; i--) {
      if(String(txData[i][1]) === String(refId)) {
        targetOpId = String(txData[i][2]).replace(/'/g, "").trim();
        msgSheet.appendRow([shadowId, new Date(), "System", targetOpId, txData[i][5], "RATING_LOG", "Rating for " + refId, "RESOLVED", "Rated", scoreDelta, tag]);
        updateMechanicScore(targetOpId);
        return { success: true };
      }
    }
    return { success: false, error: "Not found" };
  } catch(e) { return { success: false, error: e.message }; } 
  finally { lock.releaseLock(); }
}

function updateMechanicScore(opId) {
  const msgSheet = _getSheet(CONFIG.SHEETS.MESSAGING);
  const staffSheet = _getSheet(CONFIG.SHEETS.STAFF);
  const msgData = msgSheet.getDataRange().getValues();
  const staffData = staffSheet.getDataRange().getValues();
  let rawScore = 100;
  let target = String(opId).trim();
  for(let i=1; i<msgData.length; i++) {
    if(String(msgData[i][3]).trim() === target) {
      let q = Number(msgData[i][9]);
      if(!isNaN(q) && q !== 0) rawScore += q;
    }
  }
  rawScore = Math.min(100, Math.max(0, rawScore));
  for(let i=1; i<staffData.length; i++) {
    if(String(staffData[i][0]).trim() === target) {
      staffSheet.getRange(i+1, 4).setValue(rawScore);
      break;
    }
  }
}

// === ANALYTICS (ENHANCED) ===
function getAnalyticsData() {
  const txSheet = _getSheet(CONFIG.SHEETS.TRANSACTIONS);
  const staffSheet = _getSheet(CONFIG.SHEETS.STAFF);
  const msgSheet = _getSheet(CONFIG.SHEETS.MESSAGING);
  const otSheet = _getSheet(CONFIG.SHEETS.OT);

  // Build OT map (Unit -> Brand/Product)
  const otData = otSheet.getDataRange().getValues();
  const otMap = {};
  for (let i = 1; i < otData.length; i++) {
    const unitId = String(otData[i][0] || '').trim().toUpperCase();
    if (unitId) otMap[unitId] = { product: otData[i][4] || '', brand: otData[i][5] || '', semi: otData[i][2] || '' };
  }

  // Build Staff map
  const staffData = staffSheet.getDataRange().getValues();
  const staffMap = {}, mechanicsList = [];
  for (let i = 1; i < staffData.length; i++) {
    const rId = String(staffData[i][0]).replace(/'/g, "").trim();
    const rName = staffData[i][1], rScore = staffData[i][3] === "" ? 100 : Number(staffData[i][3]), rZone = staffData[i][6] || '';
    if (rId && rName) {
      staffMap[rId] = { name: rName, score: rScore, zone: rZone };
      mechanicsList.push({ id: rId, name: rName, score: rScore, zone: String(rZone) });
    }
  }

  // Build Rating map
  const msgData = msgSheet.getDataRange().getValues();
  const ratingMap = {};
  for (let i = Math.max(1, msgData.length - 500); i < msgData.length; i++) {
    if (String(msgData[i][5]) === "RATING_LOG") {
      const reqId = String(msgData[i][0]).replace("LOG-", "");
      const scoreDelta = Number(msgData[i][9]) || 0;
      const tag = String(msgData[i][10]);
      if (!ratingMap[reqId]) ratingMap[reqId] = { isRated: false, totalScore: 0, direction: null };
      ratingMap[reqId].isRated = true;
      ratingMap[reqId].totalScore += scoreDelta;
      ratingMap[reqId].direction = tag === 'Positive' ? 'up' : 'down';
    }
  }

  // Build Transactions
  const txData = txSheet.getDataRange().getValues();
  const transactions = [];
  for (let i = Math.max(1, txData.length - 100); i < txData.length; i++) {
    const row = txData[i], reqId = row[1];
    if(reqId) {
      const mId = String(row[2]).replace(/'/g, "").trim();
      const mData = staffMap[mId] || {};
      const unitInfo = String(row[5] || '');
      const primaryUnit = unitInfo.split(/[\/+]/)[0].trim().toUpperCase();
      const otInfo = otMap[primaryUnit] || {};
      const ratingInfo = ratingMap[reqId] || { isRated: false, totalScore: 0, direction: null };
      
      let reqTime = row[0], readyTime = row[9], deliveredTime = row[10];
      if (reqTime instanceof Date) reqTime = reqTime.toISOString();
      if (readyTime instanceof Date) readyTime = readyTime.toISOString();
      if (deliveredTime instanceof Date) deliveredTime = deliveredTime.toISOString();

      transactions.push({
        timestamp: reqTime, reqId: reqId, opId: mId, mechName: mData.name || row[3], mechZone: mData.zone || '',
        ot: row[4], unit: unitInfo, item: row[6], itemCount: row[7] || 1, status: row[8], notes: row[11],
        shopMins: Number(row[12]) || 0, mechMins: Number(row[13]) || 0, readyTime: readyTime, deliveredTime: deliveredTime,
        product: otInfo.product || '', brand: otInfo.brand || '', semi: otInfo.semi || '',
        isRated: ratingInfo.isRated, currentRating: ratingInfo.totalScore, ratingDirection: ratingInfo.direction
      });
    }
  }
  return { transactions: transactions.reverse(), mechanics: mechanicsList.sort((a,b) => b.score - a.score) };
}

function getDashboardNotifications(opId) {
  const sheet = _getSheet(CONFIG.SHEETS.MESSAGING);
  const data = sheet.getDataRange().getValues();
  const staffSheet = _getSheet(CONFIG.SHEETS.STAFF);
  const staffData = staffSheet.getDataRange().getValues();
  let mechMap = {};
  for(let j = 1; j < staffData.length; j++) mechMap[String(staffData[j][0]).trim()] = staffData[j][1];
  
  let notifications = [];
  for(let i = 1; i < data.length; i++) {
    const row = data[i];
    const sender = String(row[2] || '').trim(), status = String(row[7] || '').trim();
    const msgType = String(row[5] || '').trim(), reply = String(row[8] || '').trim();
    const targetId = String(row[3] || '').trim();
    
    if (sender === "Dashboard" && status === "RESOLVED" && reply && reply !== "0" && msgType !== "RATING_LOG" && (!opId || targetId === String(opId))) {
      let timeAgo = "Reciente";
      if (row[1] instanceof Date) {
        const diffMins = Math.floor((new Date() - row[1]) / 60000);
        if (diffMins < 60) timeAgo = diffMins + " min";
        else if (diffMins < 1440) timeAgo = Math.floor(diffMins / 60) + " hrs";
        else timeAgo = Math.floor(diffMins / 1440) + " días";
      }
      notifications.push({ id: row[0], timestamp: row[1] instanceof Date ? row[1].toISOString() : row[1], unit: row[4] || "General", mechId: targetId, mechName: mechMap[targetId] || "Desconocido", reply: reply, originalMsg: row[6] || "Consulta", timeAgo: timeAgo });
    }
  }
  return notifications.reverse();
}

function dismissDashboardNotification(msgId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const sheet = _getSheet(CONFIG.SHEETS.MESSAGING);
    const data = sheet.getDataRange().getValues();
    for(let i = data.length - 1; i >= 1; i--) {
      if(String(data[i][0]) === String(msgId)) { sheet.getRange(i + 1, 8).setValue("DISMISSED"); return { success: true }; }
    }
    return { success: false, error: "Not found" };
  } catch(e) { return { success: false, error: e.message }; } 
  finally { lock.releaseLock(); }
}

// === WAREHOUSE ===
function getPendingOrders() {
  var transSheet = _getSheet(CONFIG.SHEETS.TRANSACTIONS);
  var transData = transSheet.getDataRange().getValues();
  var ordersMap = {};
  for (var i = Math.max(1, transData.length - 200); i < transData.length; i++) {
    var row = transData[i], status = row[8], reqId = row[1];
    if (status === "PENDIENTE" || status === "LISTO") {
      if (!ordersMap[reqId]) ordersMap[reqId] = { reqId: reqId, timestamp: row[0] instanceof Date ? row[0].toISOString() : row[0], opInfo: row[3], otNumber: row[4], unitInfo: row[5], status: status, items: [], notes: row[11] };
      ordersMap[reqId].items.push({ name: String(row[6]), qty: row[7] });
    }
  }
  return Object.values(ordersMap).sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
}

// === ACTUALIZAR EN Code.gs ===

function getPendingOrdersEnriched() {
  const transSheet = _getSheet(CONFIG.SHEETS.TRANSACTIONS);
  const otSheet = _getSheet(CONFIG.SHEETS.OT);
  const itemSheet = _getSheet(CONFIG.SHEETS.ITEMS); // 1. Traemos DB_ITEMS
  
  // 2. Mapeo de Items (Nombre -> ID y Ubicación)
  const itemData = itemSheet.getDataRange().getValues();
  const itemMap = {};
  // i = 1 para saltar el encabezado. Col A = 0 (ID), Col B = 1 (Nombre), Col F = 5 (Ubicación)
  for (let i = 1; i < itemData.length; i++) {
    const iId = String(itemData[i][0]).trim();
    const iName = String(itemData[i][1]).trim().toUpperCase();
    const iLoc = String(itemData[i][5]).trim();
    if (iName) {
      itemMap[iName] = { id: iId, loc: iLoc || 'S/D' };
    }
  }

  // Build OT map (Unit -> Product/Brand)
  const otData = otSheet.getDataRange().getValues();
  const otMap = {};
  for (let i = 1; i < otData.length; i++) {
    const unitId = String(otData[i][CONFIG.OT_COLS.UNIT_ID] || '').trim().toUpperCase();
    if (unitId) {
      otMap[unitId] = { 
        product: otData[i][CONFIG.OT_COLS.PRODUCT] || '', 
        brand: otData[i][CONFIG.OT_COLS.BRAND] || '' 
      };
    }
  }
  
  const transData = transSheet.getDataRange().getValues();
  const ordersMap = {};
  
  for (let i = Math.max(1, transData.length - 200); i < transData.length; i++) {
    const row = transData[i], status = row[8], reqId = row[1];
    if (status === "PENDIENTE" || status === "LISTO") {
      if (!ordersMap[reqId]) {
        const unitInfo = String(row[5] || '');
        const primaryUnit = unitInfo.split(/[\/+]/)[0].trim().toUpperCase();
        const otInfo = otMap[primaryUnit] || {};
        
        ordersMap[reqId] = { 
          reqId: reqId, 
          timestamp: row[0] instanceof Date ? row[0].toISOString() : row[0], 
          opInfo: row[3], 
          otNumber: row[4], 
          unitInfo: unitInfo, 
          status: status, 
          items: [], 
          notes: row[11],
          product: otInfo.product || '',
          brand: otInfo.brand || ''
        };
      }
      
      // 3. Cruzamos el Nombre del ítem pedido con nuestro Mapa de Items
      const itemName = String(row[6]);
      const itemDetails = itemMap[itemName.toUpperCase()] || { id: '---', loc: '?' };

      ordersMap[reqId].items.push({ 
        name: itemName, 
        qty: row[7],
        id: itemDetails.id,  // Inyectamos ID
        loc: itemDetails.loc // Inyectamos Ubicación
      });
    }
  }
  return Object.values(ordersMap).sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
}

function markAsReady(reqId) { return _updateTxStatus(reqId, "LISTO", 9, 10, 13); }
function markAsDelivered(reqId) { return _updateTxStatus(reqId, "ENTREGADO", 9, 11, 14); }

// === ACTUALIZAR EN Code.gs ===



function _updateTxStatus(reqId, status, statusColIdx, timeColIdx, calcColIdx, panolOpId) {
  const sheet = _getSheet(CONFIG.SHEETS.TRANSACTIONS);
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const lock = LockService.getScriptLock();
  
  let updatedCount = 0;

  try {
    lock.waitLock(5000);
    const targetReq = String(reqId).trim();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === targetReq) {
        const rowNum = i + 1; // Fila exacta en el Excel
        
        // 1. Columna I (9) - ESTADO (PENDIENTE -> LISTO)
        sheet.getRange(rowNum, statusColIdx).setValue(status);
        
        // 2. Columna J o K (10 u 11) - TIMESTAMP
        sheet.getRange(rowNum, timeColIdx).setValue(now).setNumberFormat("dd/MM/yyyy HH:mm:ss");
        
        // 3. Columna M o N (13 o 14) - MINUTOS TRANSCURRIDOS
        let prevTime = data[i][0]; // Toma el Timestamp de creación original (Columna A)
        if (prevTime && !(prevTime instanceof Date)) prevTime = new Date(prevTime);
        let mins = prevTime ? Math.round((now.getTime() - prevTime.getTime()) / 60000) : 0;
        sheet.getRange(rowNum, calcColIdx).setValue(mins);

        // 4. Columna P (16) - FIRMA DEL PAÑOLERO
        if (panolOpId) {
          sheet.getRange(rowNum, 16).setValue(panolOpId);
        }
        
        updatedCount++;
      }
    }

    // Forzar la escritura física en la base de datos antes de liberar el sistema
    SpreadsheetApp.flush();
    return { success: updatedCount > 0, itemsUpdated: updatedCount };

  } catch(e) { 
    return { success: false, error: e.toString() }; 
  } finally { 
    lock.releaseLock(); 
  }
}

// === MECHANIC APP ===
// === ACTUALIZAR EN Code.gs ===

// === ACTUALIZAR EN Code.gs ===

function getMechanicConfig(opId) {
  const sheet = _getSheet(CONFIG.SHEETS.STAFF);
  const data = sheet.getDataRange().getValues();
  
  // Buscamos al operario coincidente
  const row = data.find(r => {
    const isIdMatch = String(r[0]).trim() === String(opId).trim(); // Columna A
    
    // Poka-Yoke: Leemos la Columna L (Índice 11) para validar el acceso.
    // Solo permitimos el ingreso si la casilla está explícitamente en TRUE.
    const hasAppAccess = r[11] === true; 
    
    return isIdMatch && hasAppAccess;
  });

  // UX Feedback: Mensaje unificado para proteger la lógica interna del sistema
  if (!row) {
    return { success: false, error: "Usuario no encontrado o sin acceso activo" };
  }

  // Nota: Mantenemos el slice(6, 11) asumiendo que tus ubicaciones/boxes 
  // siguen estando desde la Columna G (6) hasta la K (10).
  const boxes = row.slice(6, 11).map(String).filter(c => c);
  
  return { success: true, name: row[1], role: row[2], boxes: boxes };
}

function getMechanicOrders(opId) {
  const sheet = _getSheet(CONFIG.SHEETS.TRANSACTIONS);
  const data = sheet.getDataRange().getValues();
  const sId = String(opId).trim();
  const ordersMap = {};

  for (let i = Math.max(1, data.length - 300); i < data.length; i++) {
    const rowOpId = String(data[i][2]).replace(/'/g, "").trim();
    const status = String(data[i][8]);

    if (rowOpId === sId && ["PENDIENTE","LISTO","ENTREGADO","DEVOLUCION"].includes(status)) {
      const reqId = data[i][1];

      if (!ordersMap[reqId]) {
        ordersMap[reqId] = {
          reqId: reqId,
          otNumber: data[i][4],
          items: []
        };
      }
      
      // Guardamos el estado INDIVIDUAL de cada repuesto
      ordersMap[reqId].items.push({ 
        name: data[i][6], 
        qty: data[i][7],
        status: status 
      });
    }
  }
  
  // Diseño de Servicio: Determinamos el estado global del pedido basado en sus partes
  const results = Object.values(ordersMap).map(order => {
    const statuses = order.items.map(i => i.status);
    
    if (statuses.includes("PENDIENTE")) order.status = "PENDIENTE";
    else if (statuses.includes("LISTO")) order.status = "LISTO";
    else if (statuses.includes("ENTREGADO")) order.status = "ENTREGADO"; // Si queda al menos 1 entregado, permite devolver
    else order.status = "DEVOLUCION"; // Solo si TODOS están devueltos
    
    return order;
  });

  return results.reverse();
}

function submitBatchRequest(data) {
  var sheet = _getSheet(CONFIG.SHEETS.TRANSACTIONS);
  var ts = new Date();
  var reqId = "REQ-" + Math.floor(Math.random() * 1000000);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    data.items.forEach(function(i) {
      sheet.appendRow([ts, reqId, "'" + data.opId, data.mechanicName, "'" + data.otNumber, data.unitId, i.item, Number(i.qty), "PENDIENTE", "", "", i.notes || "", "", ""]);
      updateStockByName(i.item, -Number(i.qty));
    });
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; } 
  finally { lock.releaseLock(); }
}

// === REEMPLAZAR EN Code.gs ===

function findUnitOrOt(query, type) {
  const sheet = _getSheet(CONFIG.SHEETS.OT);
  const data = sheet.getDataRange().getValues();
  const q = String(query).trim().toUpperCase();

  // Poka-Yoke: Evitar búsquedas vacías o muy cortas
  if (q.length < 2) return { success: false };

  // Iteramos la Base de Datos (Omitiendo el encabezado)
  for (let i = 1; i < data.length; i++) {
    const unit = String(data[i][0]).toUpperCase().trim();     // Col A: Patente Tractor
    const ot = String(data[i][1]).trim().toUpperCase();       // Col B: OT Tractor
    const semi = String(data[i][2]).toUpperCase().trim();     // Col C: Patente Semi
    const semiOt = String(data[i][3]).trim().toUpperCase();   // Col D: OT Semi (Alternativa)

    let isMatch = false;

    // Lógica Bidireccional basada en el 'type' que envía el frontend
    if (type === 'OT') {
        // Búsqueda estricta por OT
        if (ot === q) isMatch = true;
    } else if (type === 'UNIT') {
        // Búsqueda flexible por Patente
        if (unit === q || unit.includes(q)) isMatch = true;
    } else if (type === 'SEMI_OT') {
        // Búsqueda estricta por OT de Semi
        if (semiOt === q || ot === q) isMatch = true;
    } else if (type === 'SEMI') {
        // Búsqueda flexible por Patente de Semi
        if (semi === q || semi.includes(q)) isMatch = true;
    } else {
        // Fallback: Si no hay 'type', busca en cualquier lado
        if (unit.includes(q) || semi.includes(q) || ot === q || semiOt === q) isMatch = true;
    }

    // Si encontramos coincidencia, devolvemos la Fila Completa como el Frontend espera
    if (isMatch) {
      return { 
        success: true, 
        unit: unit, 
        ot: ot, 
        semi: semi, 
        semiOt: semiOt || ot // Diseño de Servicio: Si el Semi no tiene OT propia, hereda la del Tractor
      };
    }
  }
  
  return { success: false };
}
// === REEMPLAZAR ESTA FUNCIÓN EN TU Code.gs ===

function getItemCatalog() {
  const sheet = _getSheet(CONFIG.SHEETS.ITEMS);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const itemsList = [];
  const uniqueCheck = new Set();

  // i = 1 salta el encabezado (Fila 0)
  for (let i = 1; i < data.length; i++) { 
    const name = String(data[i][1]).trim();       // Columna B (Índice 1): Nombre del repuesto
    const category = String(data[i][3]).trim();   // Columna D (Índice 3): RUBRO

    if (name) {
      // Poka-Yoke: Si el Excel no tiene rubro definido, lo agrupamos de forma segura
      const safeCategory = category ? category : "GENERAL";
      
      // Creamos una clave única para evitar enviar duplicados exactos al Frontend
      const uniqueKey = safeCategory + "|" + name;

      if (!uniqueCheck.has(uniqueKey)) {
        uniqueCheck.add(uniqueKey);
        itemsList.push({
          category: safeCategory,
          name: name
        });
      }
    } 
  }
  
  // Ordenar alfabéticamente: Primero por Categoría, luego por Nombre
  return itemsList.sort((a, b) => {
    if (a.category === b.category) {
      return a.name.localeCompare(b.name);
    }
    return a.category.localeCompare(b.category);
  });
}

function updateStockByName(itemName, qtyChange) {
  var sheet = _getSheet(CONFIG.SHEETS.ITEMS);
  var data = sheet.getDataRange().getValues();
  var search = String(itemName).trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === search) { var current = Number(data[i][4]); sheet.getRange(i + 1, 5).setValue(current + Number(qtyChange)); return true; }
  }
  return false;
}

// === ACTUALIZAR EN TU Code.gs ===

function getInventoryItems(query, categoryFilter) {
  const sheet = _getSheet(CONFIG.SHEETS.ITEMS);
  const data = sheet.getDataRange().getValues();
  const results = [];
  const searchStr = query ? String(query).toLowerCase().trim() : "";
  
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][1]); // Col B
    const category = String(data[i][3]).trim() || "GENERAL"; // Col D (Nuevo)
    
    // Si hay búsqueda, verificamos coincidencia. Si no hay, traemos todo (hasta el límite)
    if (searchStr === "" || name.toLowerCase().includes(searchStr)) {
      results.push({ 
        id: data[i][0], 
        name: name, 
        category: category, // Agregamos la categoría a la carga de datos
        brand: data[i][2], 
        stock: data[i][4], 
        loc: data[i][5] 
      });
    }
    // Límite de seguridad para no saturar la memoria del navegador móvil
    if (results.length >= 50) break;
  }
  return results;
}

function updateItemStock(id, newQty) { return _updateItemCellSafe(id, 4, newQty); }
function updateItemLoc(id, newLoc) { return _updateItemCellSafe(id, 5, newLoc); }

function _updateItemCellSafe(id, colIndex, value) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const sheet = _getSheet(CONFIG.SHEETS.ITEMS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (String(data[i][0]) === String(id)) { sheet.getRange(i + 1, colIndex + 1).setValue(value); return { success: true }; } }
    return { success: false };
  } catch (e) { return { success: false }; } 
  finally { lock.releaseLock(); }
}

const WAREHOUSE_USERS = { "1": "Ema", "6": "Matias" };
function validateWarehouseUser(key) { return WAREHOUSE_USERS[key] || null; }

// ==========================================
// 1. AGREGAR AL FINAL DE TU Code.gs
// ==========================================

// Obtiene la lista de operarios activos para el dropdown
function getStaffOptions() {
  const sheet = _getSheet(CONFIG.SHEETS.STAFF); // Tab DB_STAFF
  const data = sheet.getDataRange().getValues();
  const staff = [];
  
  // Empezamos en i = 1 para saltar los encabezados
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0]).trim();
    const name = String(data[i][1]).trim();
    if (id && name) {
      staff.push({ id: id, name: name });
    }
  }
  
  // Ordenar alfabéticamente para mejor UX
  return staff.sort((a, b) => a.name.localeCompare(b.name));
}

// Guarda la asignación en la Columna O de DB_TRANSACTIONS
function assignOperatorToOrder(reqId, opId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const sheet = _getSheet(CONFIG.SHEETS.TRANSACTIONS);
    const data = sheet.getDataRange().getValues();
    
    // Buscamos la transacción desde el final (más reciente)
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][1]) === String(reqId)) {
        // La Columna O es el índice 15 en getRange (1-based)
        sheet.getRange(i + 1, 15).setValue(opId);
        return { success: true };
      }
    }
    return { success: false, error: "Pedido no encontrado" };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}



// === 1. NUEVA FUNCIÓN: Obtener solo al personal de Pañol ===
function getPanolStaff() {
  const sheet = _getSheet(CONFIG.SHEETS.STAFF);
  const data = sheet.getDataRange().getValues();
  const panoleros = [];
  
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0]).trim(); // OP_ID
    const name = String(data[i][1]).trim(); // FULL_NAME
    const role = String(data[i][2]).trim().toUpperCase(); // ROLE
    const hasAppAccess = data[i][11] === true; // Columna L (ACCESO_APP)
    
    // Filtramos estrictamente por el rol designado para evitar listas largas
    if (id && name && hasAppAccess && (role === 'PANOL' || role === 'PAÑOL')) {
      panoleros.push({ id: id, name: name });
    }
  }
  return panoleros.sort((a, b) => a.name.localeCompare(b.name));
}

// === 2. ACTUALIZAR FUNCIONES DE ESTADO ===
// Ahora reciben un segundo parámetro: panolOpId
function markAsReady(reqId, panolOpId) { 
  return _updateTxStatus(reqId, "LISTO", 9, 10, 13, panolOpId); 
}

function markAsDelivered(reqId, panolOpId) { 
  return _updateTxStatus(reqId, "ENTREGADO", 9, 11, 14, panolOpId); 
}

// === AGREGAR AL FINAL DE Code.gs ===
// Obtiene el catálogo de unidades (Tractor y Semi) para el autocompletado
function getUnitCatalog() {
  const sheet = _getSheet(CONFIG.SHEETS.OT);
  const data = sheet.getDataRange().getValues();
  const units = new Set();
  
  // i = 1 para saltar los encabezados
  for (let i = 1; i < data.length; i++) {
    const unit = String(data[i][0]).trim().toUpperCase(); // Columna A (Unidad)
    const semi = String(data[i][2]).trim().toUpperCase(); // Columna C (Semi)
    
    if (unit) units.add(unit);
    if (semi) units.add(semi);
  }
  
  return Array.from(units).sort();
}

// === AGREGAR A Code.gs ===

function requestRefund(reqId, itemName, reason) {
  const txSheet = _getSheet(CONFIG.SHEETS.TRANSACTIONS);
  const txData = txSheet.getDataRange().getValues();
  const lock = LockService.getScriptLock();
  
  let updated = false;
  let qtyToReturn = 0;

  try {
    lock.waitLock(5000);
    
    for (let i = 1; i < txData.length; i++) {
      if (String(txData[i][1]).trim() === String(reqId).trim() && 
          String(txData[i][6]).trim() === String(itemName).trim()) {
        
        // POKA-YOKE: Prevenir doble devolución (Evita duplicar el stock matemáticamente)
        if (String(txData[i][8]).trim() === "DEVOLUCION") {
            return { success: false, error: "Este ítem ya fue devuelto previamente." };
        }

        // Ejecutamos la devolución
        txSheet.getRange(i + 1, 9).setValue("DEVOLUCION");
        const fallbackReason = reason ? reason : "Sin motivo";
        txSheet.getRange(i + 1, 15).setValue("DEVUELTO: " + fallbackReason); // Columna O
        
        qtyToReturn = Number(txData[i][7]) || 1;
        updated = true;
        break; 
      }
    }

    if (updated && qtyToReturn > 0) {
      updateStockByName(itemName, qtyToReturn); 
    }

    return { success: updated, error: updated ? null : "El repuesto exacto no fue encontrado." };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

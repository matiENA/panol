// === SINCRONIZADOR MASIVO: EXTRACCIÓN CON REFERENCIAS ABSOLUTAS ===
function SYNC_MASIVO_DB_OT_LIST() {
  
  // 1. IDENTIDAD ABSOLUTA (Poka-Yoke de enrutamiento)
  // Reemplaza "ID_DEL_EXCEL_DE_RESPUESTAS" con el ID real del archivo del Formulario (lo sacas de la URL)
  const SOURCE_SS_ID = "1HKXGsRC149Kw4aBXQwGcPVpAvObvTUFis6YV6R5cTXk"; 
  const TARGET_SS_ID = "1grLJZIYdWLRtjxK0kXobcaxj-1nZQaNnz23NU4oUDko"; // Tu DB_OT_LIST
  
  // Conectamos con los archivos físicamente separados
  const sourceSS = SpreadsheetApp.openById(SOURCE_SS_ID);
  const targetSS = SpreadsheetApp.openById(TARGET_SS_ID);

  // Conectamos con las pestañas
  const sourceSheet = sourceSS.getSheetByName("Respuestas de formulario 4");
  const targetSheet = targetSS.getSheetByName("DB_OT_LIST");

  // Validación de Cierre (Gestalt): Avisar al dev si tipeó mal un ID o Nombre
  if (!sourceSheet) throw new Error("❌ No se encontró la pestaña 'Respuestas de formulario 4' en el archivo origen.");
  if (!targetSheet) throw new Error("❌ No se encontró la pestaña 'DB_OT_LIST' en la base de datos.");

  const sourceData = sourceSheet.getDataRange().getValues();
  const latestOts = new Map();

  // 2. EXTRACCIÓN LIFO (Prägnanz: Lo último es lo válido)
  for (let i = sourceData.length - 1; i >= 1; i--) {
    const rawPlate = String(sourceData[i][4]).toUpperCase().replace(/[\s\-_.]/g, ''); // Columna E (Índice 4)
    const otNumber = sourceData[i][8]; // Columna I (Índice 8)

    if (otNumber && rawPlate) {
      // Regex para aislar patentes Mercosur/Tradicionales
      const plateRegex = /([A-Z]{2}\d{3}[A-Z]{2}|[A-Z]{3}\d{3})/g;
      const matches = rawPlate.match(plateRegex);

      if (matches) {
        matches.forEach(plate => {
          if (!latestOts.has(plate)) {
            latestOts.set(plate, otNumber);
          }
        });
      }
    }
  }

  // 3. PREPARACIÓN DEL DESTINO
  const targetData = targetSheet.getDataRange().getValues();
  const updates = [];

  for (let i = 1; i < targetData.length; i++) { 
    const dbPlateRaw = String(targetData[i][0]).toUpperCase().replace(/[\s\-_.]/g, ''); 
    const plateRegex = /([A-Z]{2}\d{3}[A-Z]{2}|[A-Z]{3}\d{3})/;
    const matchDb = dbPlateRaw.match(plateRegex);
    const cleanDbPlate = matchDb ? matchDb[0] : dbPlateRaw;

    const currentOt = targetData[i][1]; // Col B

    // Mapeo atómico
    if (latestOts.has(cleanDbPlate)) {
      updates.push([latestOts.get(cleanDbPlate)]);
    } else {
      updates.push([currentOt]);
    }
  }

  // 4. TRANSACCIÓN ATÓMICA
  if (updates.length > 0) {
    targetSheet.getRange(2, 2, updates.length, 1).setValues(updates);
  }

  console.log(`✅ Sincronización exitosa. OTs actualizadas: ${updates.length}.`);
}
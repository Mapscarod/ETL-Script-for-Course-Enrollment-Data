function llenarAntiguosC4A2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shIns = ss.getSheetByName("[Inscripción] Cursos C4 A2 Registros únicos");
  const shCursos = ss.getSheetByName("[Inscritos] Cursos C4 A2");
  const shTotales = ss.getSheetByName("[Antiguos] Totales");
  const shUnicos = ss.getSheetByName("Unicos último dato registrado");

  const datosIns = shIns.getDataRange().getValues();           // [Inscripción]
  const datosCur = shCursos.getDataRange().getValues();        // [Inscritos] (fase 1 usa largo)
  const datosDip = shTotales.getDataRange().getValues();       // [Antiguos] Totales 
  const datosUni = shUnicos.getDataRange().getValues();        // Únicos

  const FILA0 = 0;
  const ini = 1;                            // primera fila de datos
  const numFilas = Math.max(datosIns.length - 1, datosCur.length - 1);

  // Índices relevantes (como en tu código)
  const colDocumento = 1;   // A=0, B=1 ... => documento está en col A de varios orígenes 
  const colRegistrado = 13; // N

  const ANIO = 2;
  const COHORTE = 4;

  // ---------------------------
  // 1) Construir MAPS por documento para búsquedas rápidas
  // ---------------------------

  // Map Únicos: doc -> [H..L] (usar lo que necesitas)
  const mapUni = new Map();
  for (let r = ini; r < datosUni.length; r++) {
    const doc = datosUni[r][colDocumento];
    if (!doc) continue;
    mapUni.set(String(doc), {
      H: datosUni[r][7],  I: datosUni[r][8],  J: datosUni[r][9],
      K: datosUni[r][10], L: datosUni[r][11]
    });
  }

  // Map (Totales): doc -> objeto con todas las columnas que usas
  const mapDip = new Map();
  for (let r = ini; r < datosDip.length; r++) {
    const doc = datosDip[r][colDocumento];
    if (!doc) continue;
    mapDip.set(String(doc), {
      E: datosDip[r][4],
      Q: datosDip[r][16], R: datosDip[r][17], S: datosDip[r][18], T: datosDip[r][19],
      U: datosDip[r][20], V: datosDip[r][21], Z: datosDip[r][25], AA: datosDip[r][26],
      AB: datosDip[r][27], AC: datosDip[r][28], AD: datosDip[r][29], AE: datosDip[r][30],
      AF: datosDip[r][31], AG: datosDip[r][32], AH: datosDip[r][33],
      AL: datosDip[r][37], AM: datosDip[r][38], AN: datosDip[r][39], AO: datosDip[r][40],
      AP: datosDip[r][41], AQ: datosDip[r][42], AR: datosDip[r][43], AS: datosDip[r][44],
      AT: datosDip[r][45], AU: datosDip[r][46], AV: datosDip[r][47], AW: datosDip[r][48],
      AX: datosDip[r][49], AY: datosDip[r][50], AZ: datosDip[r][51],
      BB: datosDip[r][53], BC: datosDip[r][54], BD: datosDip[r][55], BE: datosDip[r][56],
      BF: datosDip[r][57], BG: datosDip[r][58], BH: datosDip[r][59]
    });
  }

  // Map Inscripción: doc1=col 16 (Q) cuando registrado, doc2=col 26 (AA) cuando NO registrado
  // Guardamos ambos en el mismo map con la misma clave de documento
  const mapIns = new Map();
  for (let r = ini; r < datosIns.length; r++) {
    const reg = datosIns[r][colRegistrado];
    const docReg = datosIns[r][16];
    const docNoReg = datosIns[r][26];
    const doc = docReg || docNoReg;
    if (!doc) continue;
    mapIns.set(String(doc), datosIns[r]);
  }

  // ---------------------------
  // 2) Fase 1: llenar B,C,D,E y H-L en [Inscritos] según Inscripción y Únicos
  //   - Construimos buffers por lotes
  // ---------------------------
  const out_B_E = [];   // columnas B..E (4 cols)
  const out_H_L = [];   // columnas H..L (5 cols)

  for (let i = 0; i < numFilas; i++) {
    const rIns = i + ini;    // fila en datosIns
    const rowIns = datosIns[rIns] || [];

    // Estado de registro y determinación de documento
    const registrado = rowIns[colRegistrado];
    const numeroDocumento = (registrado === 'Sí, soy graduado/a y confirmo que subí todos los datos necesarios para mi inscripción.'
      ? rowIns[16]   // Q
      : rowIns[26]   // AA
    ) || '';

    // B..E: [B=Documento, C=Estado, D=Nombre(B), E=Apellido(W)]
    const nombreB = rowIns[1]  || ''; // B
    const apeW    = rowIns[22] || ''; // W
    out_B_E.push([numeroDocumento, registrado || '', nombreB, apeW]);

    // H..L desde Únicos (si existe)
    const fromUni = numeroDocumento ? mapUni.get(String(numeroDocumento)) : null;
    if (fromUni) {
      out_H_L.push([fromUni.H, fromUni.I, fromUni.J, fromUni.K, fromUni.L]);
    } else {
      out_H_L.push(['', '', '', '', '']);
    }
  }

  if (out_B_E.length) {
    shCursos.getRange(2, 2, out_B_E.length, out_B_E[0].length).setValues(out_B_E); // B..E
  }
  if (out_H_L.length) {
    shCursos.getRange(2, 8, out_H_L.length, out_H_L[0].length).setValues(out_H_L); // H..L
  }

  // ⚠ IMPORTANTE: recargar datosCursos (ya con la col. B llena) para la fase 2 si deseas seguir por filas
  const datosCur2 = shCursos.getDataRange().getValues();

  // ---------------------------
  // 3) Fase 2: llenar el resto según “registrado” + fuentes (Diplomado o Inscripción)
  //   - Armamos buffers por bloques de columnas no contiguas
  // ---------------------------

  // Helper para armar una fila con valores por columnas (por claridad)
  function buildRowValores(i) {
    // i: índice base 0 para out arrays; fila real = i+2 en la hoja
    const rowCur = datosCur2[i + 1] || [];
    const estado = rowCur[2];                 // C en [Inscritos]
    const doc    = rowCur[colDocumento] || ''; // B en [Inscritos] (ya rellenado en fase 1)

    let col = {
      A:'', E:'', F:'', G:'',
      O:'', P:'', Q:'', R:'', S:'', T:'', U:'', V:'',
      X:'', Z:'', AA:'', AB:'', AC:'', AD:'', AE:'', AF:'', AG:'', AH:'',
      AL:'', AM:'', AN:'', AO:'', AP:'', AQ:'', AR:'', AS:'', AT:'', AU:'', AV:'',
      AW:'', AX:'', AY:'', AZ:'', BA:'', BB:'', BC:'', BD:'', BE:'', BF:'', BG:'',
      BI:'', BJ:'', BK:'', BL:''
      // (agregar si se necesitan más)
    };

    // Datos desde Inscripción (si existe registro)
    const rowIns = doc ? mapIns.get(String(doc)) : null;

    if (estado === 'Sí, soy graduado/a y confirmo que subí todos los datos necesarios para mi inscripción.') {
      // Caso "registrado": se combina Diplomado + algunos campos de Inscripción
      const dDip = mapDip.get(String(doc));
      if (dDip) {
        col.E = dDip.E;
        col.F = ANIO;
        col.G = COHORTE;
        col.Q = dDip.Q; col.R = dDip.R; col.S = dDip.S; col.T = dDip.T;
        col.U = dDip.U; col.V = dDip.V; col.Z = dDip.Z; col.AA = dDip.AA;
        col.AB = dDip.AB; col.AC = dDip.AC; col.AD = dDip.AD; col.AE = dDip.AE;
        col.AF = dDip.AF; col.AG = dDip.AG; col.AH = dDip.AH; col.AL = dDip.AL;
        col.AM = dDip.AM; col.AN = dDip.AN; col.AO = dDip.AO; col.AP = dDip.AP;
        col.AQ = dDip.AQ; col.AR = dDip.AR; col.AS = dDip.AS; col.AT = dDip.AT;
        col.AU = dDip.AU; col.AV = dDip.AV; col.AW = dDip.AW; col.AX = dDip.AX;
        col.AY = dDip.AY; col.AZ = dDip.AZ; col.BB = dDip.BB; col.BC = dDip.BC;
        col.BD = dDip.BD; col.BE = dDip.BE; col.BF = dDip.BF; col.BG = dDip.BG;
      }
      if (rowIns) {
        col.A  = rowIns[0];
        col.O  = rowIns[14];
        col.P  = rowIns[15];
        col.R  = rowIns[17];       // sobreescribe la R del diplomado si así lo definiste
        col.AS = rowIns[57];
        col.S  = rowIns[18];
        col.X  = rowIns[2];
        col.BA = rowIns[67];
        col.BI = rowIns[3];
        col.BJ = rowIns[5];
        col.BK = rowIns[7];
        col.BL = rowIns[9];
        // col.BM, BN si los necesitas
      }
    } else {
      // Caso “NO registrado”: sólo Inscripción
      if (rowIns) {
        col.A  = rowIns[0];
        col.F  = ANIO;
        col.G  = COHORTE;
        col.E  = rowIns[22];
        col.O  = rowIns[23];
        col.P  = rowIns[24];
        col.Q  = rowIns[42];
        col.R  = rowIns[41];
        col.S  = rowIns[41];
        col.T  = rowIns[25];
        col.U  = rowIns[28];
        col.V  = rowIns[29];
        col.X  = rowIns[2];
        col.Z  = rowIns[30];
        col.AA = rowIns[31];
        col.AB = rowIns[32];
        col.AC = rowIns[33];
        col.AD = rowIns[34];
        col.AE = rowIns[35];
        col.AF = rowIns[56];
        col.AG = rowIns[45];
        col.AH = rowIns[46];
        col.AL = rowIns[36];
        col.AM = rowIns[37];
        col.AN = rowIns[38];
        col.AO = rowIns[39];
        col.AP = rowIns[40];
        col.AQ = rowIns[43];
        col.AR = rowIns[44];
        col.AS = rowIns[58];
        col.AT = rowIns[47];
        col.AU = rowIns[48];
        col.AV = rowIns[49];
        col.AW = rowIns[50];
        col.AX = rowIns[51];
        col.AY = rowIns[53];
        col.AZ = rowIns[54];
        col.BA = rowIns[67];
        col.BE = rowIns[52];
        col.BF = rowIns[21];
        col.BG = rowIns[20];
        col.BI = rowIns[3];
        col.BJ = rowIns[5];
        col.BK = rowIns[7];
        col.BL = rowIns[9];
      }
    }
    return col;
  }

  // Buffers por bloques (agrupa columnas contiguas para setValues)
  const blkA_A   = []; // A
  const blkE_G   = []; // E..G
  const blkO_V   = []; // O..V
  const blkX_X   = []; // X
  const blkZ_AE  = []; // Z..AE
  const blkAF_AH = []; // AF..AH
  const blkAL_AR = []; // AL..AR
  const blkAS_AZ = []; // AS..AZ
  const blkBA_BG = []; // BA..BG
  const blkBI_BL = []; // BI..BL

  for (let i = 0; i < numFilas; i++) {
    const c = buildRowValores(i);

    blkA_A.push([c.A]);
    blkE_G.push([c.E, c.F, c.G]);
    blkO_V.push([c.O, c.P, c.Q, c.R, c.S, c.T, c.U, c.V]);
    blkX_X.push([c.X]);
    blkZ_AE.push([c.Z, c.AA, c.AB, c.AC, c.AD, c.AE]);
    blkAF_AH.push([c.AF, c.AG, c.AH]);
    blkAL_AR.push([c.AL, c.AM, c.AN, c.AO, c.AP, c.AQ, c.AR]);
    blkAS_AZ.push([c.AS, c.AT, c.AU, c.AV, c.AW, c.AX, c.AY, c.AZ]);
    blkBA_BG.push([c.BA, c.BB || '', c.BC || '', c.BD || '', c.BE || '', c.BF || '', c.BG || '']);
    blkBI_BL.push([c.BI, c.BJ, c.BK, c.BL]);
  }

  const startRow = 2;

  if (numFilas > 0) {
    shCursos.getRange(startRow, 1, numFilas, 1).setValues(blkA_A);          // A
    shCursos.getRange(startRow, 5, numFilas, 3).setValues(blkE_G);          // E..G
    shCursos.getRange(startRow, 15, numFilas, 8).setValues(blkO_V);         // O..V
    shCursos.getRange(startRow, 24, numFilas, 1).setValues(blkX_X);         // X
    shCursos.getRange(startRow, 26, numFilas, 6).setValues(blkZ_AE);        // Z..AE
    shCursos.getRange(startRow, 32, numFilas, 3).setValues(blkAF_AH);       // AF..AH
    shCursos.getRange(startRow, 38, numFilas, 7).setValues(blkAL_AR);       // AL..AR
    shCursos.getRange(startRow, 45, numFilas, 8).setValues(blkAS_AZ);       // AS..AZ
    shCursos.getRange(startRow, 53, numFilas, 7).setValues(blkBA_BG);       // BA..BG
    shCursos.getRange(startRow, 61, numFilas, 4).setValues(blkBI_BL);       // BI..BL
  }

  // 4) Edad (col W) por lote
  calcularEdadC4A2__mejor();
  SpreadsheetApp.flush();
  Utilities.sleep(1000);   

  extraerLinkCarpeta();
  SpreadsheetApp.flush();
  Utilities.sleep(1000);

  actualizarColumnaCC4A2();
  SpreadsheetApp.flush();
  Utilities.sleep(1000);
  
  filtrarYExportarNoAprobadosPorTerritorioC4();
  
}


// --- Cálculo de edad tomando la fecha de nacimiento desde la columna V y escribiendo en W ---
function calcularEdadC4A2__mejor() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shCursos = ss.getSheetByName("[Inscritos] Cursos C4 A2");
  const data = shCursos.getDataRange().getValues();

  // V -> índice 21 (0-based). W -> índice 22 (0-based).
  let idxFNac = 21; // V
  let idxEdad = 22; // W

  // Alternativa robusta por encabezado (descomenta si lo quieres dinámico):
  // const headers = data[0].map(String);
  // idxFNac = headers.indexOf("¿Cuál es tu fecha de nacimiento?") !== -1 ? headers.indexOf("¿Cuál es tu fecha de nacimiento?") : 21;
  // idxEdad = headers.indexOf("Edad") !== -1 ? headers.indexOf("Edad") : 22;

  const out = [];
  const fechaCorte = new Date(2025, 10, 29); // fecha de cierre del curso

  for (let r = 1; r < data.length; r++) {
    const valor = data[r][idxFNac];
    let edad = '';
    const fechaNac = _toDateFlexibleC4(valor);
    if (fechaNac) {
      edad = _edadDesdeFechaC4(fechaNac, fechaCorte);
      if (edad < 0 || isNaN(edad)) edad = ''; // casos raros o fechas futuras
    }
    out.push([edad]);
  }

  if (out.length) {
    shCursos.getRange(2, idxEdad + 1, out.length, 1).setValues(out); // Escribe en W
  }
}

// Convierte valor de celda a Date, soportando Date, número serial de Sheets y texto "dd/mm/aaaa"
function _toDateFlexibleC4(v) {
  if (!v || typeof v === 'string' && v.toString().toUpperCase().includes('#NUM')) return null;

  if (v instanceof Date && !isNaN(v)) return v;

  // Número serial de Sheets (si la celda no está formateada como fecha al leerla)
  if (typeof v === 'number') {
    // Epoch de Sheets: 1899-12-30
    const ms = (v - 25569) * 86400 * 1000;
    const d = new Date(ms);
    return isNaN(d) ? null : d;
  }

  // Texto "dd/mm/aaaa" o con guiones
  if (typeof v === 'string') {
    const m = v.trim().match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (m) {
      let d = parseInt(m[1], 10);
      let mo = parseInt(m[2], 10) - 1;
      let y = parseInt(m[3], 10);
      if (y < 100) y += 2000; // por si llega "09", "03", etc.
      const fecha = new Date(y, mo, d);
      return isNaN(fecha) ? null : fecha;
    }
    // Último intento: confiar en Date()
    const f = new Date(v);
    return isNaN(f) ? null : f;
  }

  return null;
}

function _edadDesdeFechaC4(nacimiento, fechaCorte) {
  let edad = fechaCorte.getFullYear() - nacimiento.getFullYear();
  const m = fechaCorte.getMonth() - nacimiento.getMonth();
  if (m < 0 || (m === 0 && fechaCorte.getDate() < nacimiento.getDate())) edad--;
  return edad;
}

function actualizarColumnaCC4A2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const unicosSheet = ss.getSheetByName("[Inscritos] Cursos C4 A2");

  const lastRow = unicosSheet.getLastRow();
  if (lastRow < 2) return;

  const rangoDatos = unicosSheet.getRange(2, 8, lastRow - 1, 5).getValues(); // H:L
  const colC = unicosSheet.getRange(2, 3, lastRow - 1, 1); // C

  const nuevasCeldas = rangoDatos.map(row => {
    const contieneAprobado = row.some(valor => valor && String(valor).includes("Aprobado"));
    return contieneAprobado
      ? ["Sí, soy graduado/a y confirmo que subí todos los datos necesarios para mi inscripción."]
      : ["No cuento con certificado de cursos o diplomado del proyecto Jóvenes 4.0"];
  });

  colC.setValues(nuevasCeldas);
  console.log("Columna C actualizada correctamente en función de H:L.");
}
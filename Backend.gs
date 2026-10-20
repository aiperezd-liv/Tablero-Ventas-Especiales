/**
 * Función backend principal llamada desde index.html
 */
function getDataFromSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "Ventas_Especiales_Acumualdo";
    var sheet = ss.getSheetByName(sheetName);
    
    // Si no encuentra el nombre exacto, busca una pestaña similar o la activa
    if (!sheet) {
      var sheets = ss.getSheets();
      for (var i = 0; i < sheets.length; i++) {
        if (sheets[i].getName().toLowerCase().indexOf("ventas_especiales") !== -1) {
          sheet = sheets[i];
          break;
        }
      }
      if (!sheet) sheet = ss.getActiveSheet();
    }

    var values = [];

    // MECANISMO DE LECTURA COMPATIBLE CON DATASOURCE (CONNECTED SHEETS) Y HOJAS REGULARES
    if (sheet.getType() === SpreadsheetApp.SheetType.DATASOURCE) {
      var dsTables = sheet.getDataSourceTables();
      if (dsTables && dsTables.length > 0) {
        // Garantizar que la tabla esté actualizada y obtener el rango compilado
        values = dsTables[0].getDataRange().getDisplayValues();
      } else {
        // Lectura directa desde el Spreadsheet por rango A:AM
        var maxRows = sheet.getMaxRows() || 5000;
        values = sheet.getRange(1, 1, maxRows, 39).getDisplayValues();
      }
    } else {
      // Hoja estándar de Google Sheets
      values = sheet.getDataRange().getValues();
    }

    if (!values || values.length <= 1) {
      throw new Error("La hoja '" + sheet.getName() + "' no contiene filas de datos.");
    }

    // Encabezados en primera fila
    var headers = values[0].map(function(h) { return String(h).trim(); });
    
    // Helper para obtener índice de columna por nombre (sin distinción de mayúsculas/espacios)
    function getColIdx(colName) {
      var target = colName.toLowerCase().replace(/_/g, "");
      for (var j = 0; j < headers.length; j++) {
        var cleanHeader = headers[j].toLowerCase().replace(/_/g, "");
        if (cleanHeader === target) return j;
      }
      return -1;
    }

    // Mapeo de índices de columna
    var idx = {
      ANIO: getColIdx("ANIO"),
      Trimestre: getColIdx("Trimestre"),
      Producto: getColIdx("Producto"),
      Segmento: getColIdx("Segmento"),
      Temp_Venta: getColIdx("Temp_Venta"),
      Tipo_de_Venta: getColIdx("Tipo_de_Venta"),
      Solicitudes: getColIdx("Solicitudes"),
      Cant_Tiendas: getColIdx("Cant_Tiendas"),
      Aprobadas: getColIdx("Aprobadas"),
      Cuentas: getColIdx("Cuentas"),
      CuentasNHTF: getColIdx("CuentasNHTF"),
      SumHVH: getColIdx("SumHVH"),
      SumScore: getColIdx("SumScore"),
      SumEdad: getColIdx("SumEdad"),
      SumFemenino: getColIdx("SumFemenino"),
      SumLicPos: getColIdx("SumLicPos"),
      SumIngreso: getColIdx("SumIngreso"),
      SUM_PTI: getColIdx("SUM_PTI"),
      SUM_BTI: getColIdx("SUM_BTI"),
      SUM_LINEA_ONUS: getColIdx("SUM_LINEA_ONUS"),
      SUM_LINEA_OFFUS: getColIdx("SUM_LINEA_OFFUS"),
      SUM_IQ_NUM_CONS_3M: getColIdx("SUM_IQ_NUM_CONS_3M"),
      SUM_IQ_NUM_CONS_6M: getColIdx("SUM_IQ_NUM_CONS_6M"),
      N_TARJETAS: getColIdx("N_TARJETAS"),
      SUM_UTILIZACION: getColIdx("SUM_UTILIZACION"),
      SUM_LTI_TOTAL: getColIdx("SUM_LTI_TOTAL")
    };

    // Función auxiliar para parsear valores numéricos limpios
    function parseNum(val) {
      if (val === null || val === undefined || val === "" || val === "null" || val === "NULL") return 0;
      if (typeof val === "number") return val;
      var clean = String(val).replace(/,/g, "").replace(/\$/g, "").replace(/%/g, "").trim();
      var num = parseFloat(clean);
      return isNaN(num) ? 0 : num;
    }

    var records = [];

    // Procesar filas (a partir de la fila 1)
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      if (!row || row.length === 0 || row[0] === "") continue;

      var tempVentaVal = String(row[idx.Temp_Venta] || "").trim();
      
      // Mapeo automático de Mes a partir de Temp_Venta / Campaña
      var mesCalculado = "Otros";
      var tempLower = tempVentaVal.toLowerCase();
      if (tempLower.indexOf("dec") !== -1 || tempLower.indexOf("dic") !== -1) mesCalculado = "Diciembre";
      else if (tempLower.indexOf("nov") !== -1 || tempLower.indexOf("black") !== -1) mesCalculado = "Noviembre";
      else if (tempLower.indexOf("oct") !== -1) mesCalculado = "Octubre";
      else if (tempLower.indexOf("jun") !== -1) mesCalculado = "Junio";
      else if (tempLower.indexOf("may") !== -1 || tempLower.indexOf("mamas") !== -1 || tempLower.indexOf("papas") !== -1) mesCalculado = "Mayo";
      else if (tempLower.indexOf("feb") !== -1) mesCalculado = "Febrero";
      else if (tempLower.indexOf("abr") !== -1) mesCalculado = "Abril";
      else if (tempLower.indexOf("ene") !== -1 || tempLower.indexOf("jan") !== -1) mesCalculado = "Enero";
      else if (tempLower.indexOf("bau") !== -1) mesCalculado = "BAU";

      records.push({
        anio: String(row[idx.ANIO] || "").trim(),
        trimestre: String(row[idx.Trimestre] || "").trim(),
        producto: String(row[idx.Producto] || "").trim(),
        segmento: String(row[idx.Segmento] || "").trim(),
        tempVenta: tempVentaVal,
        tipoVenta: String(row[idx.Tipo_de_Venta] || "").trim(),
        mes: mesCalculado,
        
        // Métricas numéricas
        solicitudes: parseNum(row[idx.Solicitudes]),
        cantTiendas: parseNum(row[idx.Cant_Tiendas]),
        aprobadas: parseNum(row[idx.Aprobadas]),
        cuentas: parseNum(row[idx.Cuentas]),
        cuentasNHTF: parseNum(row[idx.CuentasNHTF]),
        sumHVH: parseNum(row[idx.SumHVH]),
        sumScore: parseNum(row[idx.SumScore]),
        sumEdad: parseNum(row[idx.SumEdad]),
        sumFemenino: parseNum(row[idx.SumFemenino]),
        sumLicPos: parseNum(row[idx.SumLicPos]),
        sumIngreso: parseNum(row[idx.SumIngreso]),
        sumPTI: parseNum(row[idx.SUM_PTI]),
        sumBTI: parseNum(row[idx.SUM_BTI]),
        sumLineaOnUs: parseNum(row[idx.SUM_LINEA_ONUS]),
        sumLineaOffUs: parseNum(row[idx.SUM_LINEA_OFFUS]),
        sumCons3M: parseNum(row[idx.SUM_IQ_NUM_CONS_3M]),
        sumCons6M: parseNum(row[idx.SUM_IQ_NUM_CONS_6M]),
        nTarjetas: parseNum(row[idx.N_TARJETAS]),
        sumUtilizacion: parseNum(row[idx.SUM_UTILIZACION]),
        sumLtiTotal: parseNum(row[idx.SUM_LTI_TOTAL])
      });
    }

    return {
      status: "success",
      totalRecords: records.length,
      data: records
    };

  } catch (err) {
    return {
      status: "error",
      message: err.toString()
    };
  }
}

/**
 * Servir la aplicación web HTML
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Suburbia - Ventas Especiales')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

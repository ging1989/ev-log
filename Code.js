const SHEET_ID = '1OSQCOL1HYH1QOchPPHj5GwBX5NqVKykwQoUgvUVq14Y';

const SHEET_NAMES = {
  EV_LOG: 'EV_Log',
  STATIONS: 'stations'
};

/**
 * GET API Router
 *
 * Example:
 * ?action=getData
 * ?action=getStations
 */
function doGet(e) {

  try {

    const action = e.parameter.action;

    switch (action) {

      case 'getData':
        return getData();

      case 'getStations':
        return getStations();

      default:
        return jsonResponse({
          success: false,
          message: 'Invalid action'
        }, 400);
    }

  } catch (error) {

    return jsonResponse({
      success: false,
      message: error.toString()
    }, 500);
  }
}

/**
 * POST API
 * Save charging log
 */
function doPost(e) {

  try {

    if (!e || !e.postData || !e.postData.contents) {

      return jsonResponse({
        success: false,
        message: 'No request body found'
      }, 400);
    }

    const body = JSON.parse(e.postData.contents);

    validatePayload(body);

    const sheet = getSheet(SHEET_NAMES.EV_LOG);

    const id = Utilities.getUuid();
    const createdAt = new Date();

    sheet.appendRow([
      id,
      createdAt,
      body.date || '',
      body.station || '',
      body.trip || '',
      Number(body.priceBeforeDiscount) || 0,
      Number(body.kwh) || 0,
      Number(body.discount) || 0,
      Number(body.finalPrice) || 0,
      Number(body.bahtPerKwh) || 0
    ]);

    return jsonResponse({
      success: true,
      message: 'Charging log saved',
      data: {
        id: id
      }
    });

  } catch (error) {

    return jsonResponse({
      success: false,
      message: error.message || error.toString()
    }, 500);
  }
}

/**
 * Get all charging logs
 */
function getData() {

  const sheet = getSheet(SHEET_NAMES.EV_LOG);

  const rows = sheet.getDataRange().getValues();

  if (rows.length <= 1) {
    return jsonResponse({
      success: true,
      data: []
    });
  }

  const data = [];

  for (let i = 1; i < rows.length; i++) {

    const row = rows[i];

    if (!row[0]) continue;

    data.push({
      id: row[0],

      createdAt: row[1]
        ? Utilities.formatDate(
            new Date(row[1]),
            Session.getScriptTimeZone(),
            'yyyy-MM-dd HH:mm:ss'
          )
        : '',

      date: row[2]
        ? Utilities.formatDate(
            new Date(row[2]),
            Session.getScriptTimeZone(),
            'yyyy-MM-dd'
          )
        : '',

      station: row[3] || '',
      trip: row[4] || '',

      priceBeforeDiscount: Number(row[5]) || 0,
      kwh: Number(row[6]) || 0,
      discount: Number(row[7]) || 0,
      finalPrice: Number(row[8]) || 0,
      bahtPerKwh: Number(row[9]) || 0
    });
  }

  data.sort(function(a, b) {
    return new Date(b.date) - new Date(a.date);
  });

  return jsonResponse({
    success: true,
    data: data
  });
}

/**
 * Get charging stations
 */
function getStations() {

  const sheet = getSheet(SHEET_NAMES.STATIONS);

  const rows = sheet.getDataRange().getValues();

  const stations = [];

  for (let i = 1; i < rows.length; i++) {

    const name = String(rows[i][0] || '').trim();

    if (!name) continue;

    stations.push({
      name: name,
      onPeak: Number(rows[i][1]) || null,
      offPeak: Number(rows[i][2]) || null,
      singlePrice: Number(rows[i][3]) || null
    });
  }

  stations.sort(function(a, b) {
    return a.name.localeCompare(b.name);
  });

  return jsonResponse({
    success: true,
    data: stations
  });
}

/**
 * Validate request payload
 */
function validatePayload(body) {

  if (!body) {
    throw new Error('Request body is required');
  }

  if (!body.date) {
    throw new Error('date is required');
  }

  if (!body.station) {
    throw new Error('station is required');
  }

  if (Number(body.kwh) <= 0) {
    throw new Error('kwh must be greater than 0');
  }
}

/**
 * Open sheet helper
 */
function getSheet(sheetName) {

  const ss = SpreadsheetApp.openById(SHEET_ID);

  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`Sheet not found: ${sheetName}`);
  }

  return sheet;
}

/**
 * JSON response helper
 */
function jsonResponse(data) {

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
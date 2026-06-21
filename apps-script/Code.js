const CHECKINS_SHEET_NAME = 'Checkins';
const USER_DAILY_STATUS_SHEET_NAME = 'UserDailyStatus';
const LINE_LOGIN_CHANNEL_ID = '2010462217';

const BLESSING_COLUMNS = {
  morningPrayer: 4,
  prayerMeeting: 5,
  smallGroup: 6,
  sunday: 7,
  trainingSystem: 8
};

function doGet(e) {
  return jsonResponse({
    success: true,
    message: 'Apps Script API is working'
  });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const idToken = body.idToken;
    const blessingType = body.blessingType;

    if (!idToken) {
      return jsonResponse({
        success: false,
        error: 'Missing idToken'
      });
    }

    if (!blessingType || !BLESSING_COLUMNS[blessingType]) {
      return jsonResponse({
        success: false,
        error: 'Invalid blessingType: ' + blessingType
      });
    }

    const lineProfile = verifyLineIdToken(idToken);

    if (!lineProfile || !lineProfile.sub) {
      return jsonResponse({
        success: false,
        error: 'Invalid LINE token'
      });
    }

    const now = new Date();
    const today = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd');

    const lock = LockService.getScriptLock();

    try {
      lock.waitLock(5000);

      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

      appendCheckinLog({
        spreadsheet,
        date: today,
        lineUserId: lineProfile.sub,
        displayName: lineProfile.name || '',
        blessingType,
        now
      });

      updateUserDailyStatus({
        spreadsheet,
        date: today,
        lineUserId: lineProfile.sub,
        displayName: lineProfile.name || '',
        blessingType,
        now
      });

    } finally {
      lock.releaseLock();
    }

    return jsonResponse({
      success: true,
      date: today,
      lineUserId: lineProfile.sub,
      displayName: lineProfile.name || '',
      blessingType
    });

  } catch (error) {
    return jsonResponse({
      success: false,
      error: String(error)
    });
  }
}

function appendCheckinLog({ spreadsheet, date, lineUserId, displayName, blessingType, now }) {
  const sheet = spreadsheet.getSheetByName(CHECKINS_SHEET_NAME);

  if (!sheet) {
    throw new Error('Sheet not found: ' + CHECKINS_SHEET_NAME);
  }

  sheet.appendRow([
    now,
    lineUserId,
    displayName,
    blessingType,
    date
  ]);
}

function updateUserDailyStatus({ spreadsheet, date, lineUserId, displayName, blessingType, now }) {
  const sheet = spreadsheet.getSheetByName(USER_DAILY_STATUS_SHEET_NAME);

  if (!sheet) {
    throw new Error('Sheet not found: ' + USER_DAILY_STATUS_SHEET_NAME);
  }

  const targetColumn = BLESSING_COLUMNS[blessingType];
  const lastRow = sheet.getLastRow();

  let targetRow = -1;

  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

    for (let i = 0; i < values.length; i++) {
      const rowDate = String(values[i][0]);
      const rowLineUserId = String(values[i][1]);

      if (rowDate === date && rowLineUserId === lineUserId) {
        targetRow = i + 2;
        break;
      }
    }
  }

  if (targetRow === -1) {
    sheet.appendRow([
      date,
      lineUserId,
      displayName,
      false,
      false,
      false,
      false,
      false,
      now
    ]);

    targetRow = sheet.getLastRow();
  }

  sheet.getRange(targetRow, 3).setValue(displayName);
  sheet.getRange(targetRow, targetColumn).setValue(true);
  sheet.getRange(targetRow, 9).setValue(now);
}

function verifyLineIdToken(idToken) {
  const response = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post',
    payload: {
      id_token: idToken,
      client_id: LINE_LOGIN_CHANNEL_ID
    },
    muteHttpExceptions: true
  });

  const statusCode = response.getResponseCode();
  const text = response.getContentText();

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('LINE verify failed: ' + text);
  }

  return JSON.parse(text);
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

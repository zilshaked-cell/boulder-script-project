// המרת מספר עמודה (1-מבוסס) לאות בגיליון (A, B, ... AA)
function EMP_columnNumberToLetter_(col) {
  let temp = col;
  let letter = '';
  while (temp > 0) {
    const rem = (temp - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    temp = Math.floor((temp - 1) / 26);
  }
  return letter;
}

// בודקת אם כותרות הכרטיסייה "פרטי עובדים" השתנו מאז הפעם הקודמת.
// אם כן — שולחת אליך מייל עם פירוט + פרומפט מוכן ל-GPT.
function EMP_checkEmployeesHeader_() {
  const SHEET_NAME = 'פרטי עובדים';
  const PROPS_KEY = 'EMP_HEADERS_' + SHEET_NAME;

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    // אם הכרטיסייה נמחקה לגמרי — אין מה לעשות פה.
    return;
  }

  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;

  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const currentHeaders = headerRow.map(v => (v || '').toString().trim());

  const props = PropertiesService.getScriptProperties();
  const prevJson = props.getProperty(PROPS_KEY);

  // ריצה ראשונה – רק שומר מבנה, לא שולח מייל
  if (!prevJson) {
    props.setProperty(PROPS_KEY, JSON.stringify(currentHeaders));
    return;
  }

  /** @type {string[]} */
  const prevHeaders = JSON.parse(prevJson);

  // אם הכל זהה – אין שינוי, יוצאים
  if (prevHeaders.length === currentHeaders.length &&
      prevHeaders.every((h, i) => h === currentHeaders[i])) {
    return;
  }

  // בונים תיאור שינויים לפי עמודות
  const changes = [];
  const maxLen = Math.max(prevHeaders.length, currentHeaders.length);

  for (let i = 0; i < maxLen; i++) {
    const oldVal = (prevHeaders[i] || '').toString();
    const newVal = (currentHeaders[i] || '').toString();
    if (oldVal === newVal) continue;

    const colLetter = EMP_columnNumberToLetter_(i + 1);
    if (!oldVal && newVal) {
      changes.push('עמודה ' + colLetter + ': נוספה כותרת חדשה "' + newVal + '".');
    } else if (oldVal && !newVal) {
      changes.push('עמודה ' + colLetter + ': הכותרת "' + oldVal + '" נמחקה.');
    } else {
      changes.push(
        'עמודה ' + colLetter + ': הכותרת שונתה מ-"' + oldVal + '" ל-"' + newVal + '".'
      );
    }
  }

  const subject = 'שינוי במבנה הכותרות של "פרטי עובדים"';
  const bodyLines = [];

  bodyLines.push('זוהה שינוי בשורת הכותרות (שורה 1) בגיליון "פרטי עובדים".');
  bodyLines.push('');
  bodyLines.push('שינויים מזוהים:');
  bodyLines.push(changes.length ? changes.join('\n') : '(השינוי הוא באורך/מיקום בלבד).');
  bodyLines.push('');
  bodyLines.push('כותרות קודמות:');
  bodyLines.push(prevHeaders.join(' | '));
  bodyLines.push('');
  bodyLines.push('כותרות נוכחיות:');
  bodyLines.push(currentHeaders.join(' | '));
  bodyLines.push('');
  bodyLines.push('פרומפט מוכן ל-GPT (להעתקה):');
  bodyLines.push('');
  bodyLines.push('```');
  bodyLines.push('אני שקד. בגיליון ניהול העובדים של בולדר חיפה השתנו כותרות העמודות בכרטיסייה "פרטי עובדים".');
  bodyLines.push('הכותרות הישנות היו:');
  bodyLines.push(prevHeaders.join(' | '));
  bodyLines.push('הכותרות החדשות הן:');
  bodyLines.push(currentHeaders.join(' | '));
  bodyLines.push('');
  bodyLines.push('אני עובד בשיחה שנקראת "תכנון מערכת ID\'S".');
  bodyLines.push('תעזור לעדכן את הסקריפטים כך שיעבדו עם המבנה החדש בלי לשנות את מבנה הטבלה.');
  bodyLines.push('```');

  MailApp.sendEmail({
    to: 'ZIL.SHAKED@GMAIL.COM',
    subject: subject,
    body: bodyLines.join('\n')
  });

  // מעדכן את המצב האחרון כדי שלא תקבל מיילים כפולים על אותו שינוי
  props.setProperty(PROPS_KEY, JSON.stringify(currentHeaders));
}

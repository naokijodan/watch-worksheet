/**
 * Watch Worksheet Generator - 統合版
 * - FedEx新形式対応（横型レイアウト、複数時計対応：Watch 1～3）
 * - DHL形式対応（縦型レイアウト、2パターン）
 * - 新項目追加：Primary function, Battery origin, Movement size, Quantity
 * - FedExフォーマットのデータで両方に対応
 */

// ===========================================
// メニュー
// ===========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Watch Worksheet')
    .addItem('初期設定', 'showSettings')
    .addItem('新しいワークシート作成', 'showMainInput')
    .addSeparator()
    .addItem('新形式(FedEx)でPDF出力', 'exportCurrentSheetToNewFormatPDF')
    .addItem('DHL形式でPDF出力', 'exportCurrentSheetToDHLPDF')
    .addItem('DHL用ver.2でPDF出力', 'exportCurrentSheetToDHLPDFv2')
    .addItem('PDF出力用フォルダ設定', 'showFolderSettings')
    .addSeparator()
    .addItem('設定リセット', 'resetSettings')
    .addToUi();
}

// ===========================================
// 設定UI／保存
// ===========================================

/** 初期設定画面 */
function showSettings() {
  const html = HtmlService.createTemplateFromFile('settings');
  const config = getCompanyConfig();
  html.companyName = config.companyName || '';
  html.nameAndTitle = config.nameAndTitle || '';
  html.email = config.email || '';
  const htmlOutput = html.evaluate().setWidth(500).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '初期設定 - 会社情報');
}

/** メイン入力（ChatGPTテキスト→ワークシート作成） */
function showMainInput() {
  const config = getCompanyConfig();
  if (!config.companyName || !config.nameAndTitle || !config.email) {
    SpreadsheetApp.getUi().alert(
      'エラー',
      '初期設定が完了していません。\n「Watch Worksheet > 初期設定」から会社情報を入力してください。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  const html = HtmlService.createHtmlOutputFromFile('input').setWidth(700).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Watch Worksheet 作成');
}

/** 設定リセット */
function resetSettings() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('確認', '設定をリセットしますか？', ui.ButtonSet.YES_NO);
  if (response === ui.Button.YES) {
    PropertiesService.getScriptProperties().deleteProperty('COMPANY_CONFIG');
    ui.alert('完了', '設定をリセットしました。', ui.ButtonSet.OK);
  }
}

/** 会社設定保存（HTML側からgoogle.script.runで呼ぶ） */
function saveCompanyConfig(companyName, nameAndTitle, email) {
  try {
    const config = { companyName, nameAndTitle, email };
    PropertiesService.getScriptProperties().setProperty('COMPANY_CONFIG', JSON.stringify(config));
    return { success: true, message: '設定を保存しました。' };
  } catch (error) {
    console.error('設定保存エラー:', error);
    return { success: false, message: '設定の保存に失敗しました: ' + error.message };
  }
}

/** 会社設定取得 */
function getCompanyConfig() {
  try {
    const json = PropertiesService.getScriptProperties().getProperty('COMPANY_CONFIG');
    if (!json) return { companyName: '', nameAndTitle: '', email: '' };
    return JSON.parse(json);
  } catch (e) {
    console.error('設定取得エラー:', e);
    return { companyName: '', nameAndTitle: '', email: '' };
  }
}

// ===========================================
// PDFフォルダ設定
// ===========================================

function showFolderSettings() {
  const ui = SpreadsheetApp.getUi();
  const currentFolder = getPDFFolder();
  const response = ui.prompt(
    'PDF出力フォルダ設定',
    `現在のフォルダ: ${currentFolder ? currentFolder.getName() : 'ルートフォルダ'}\n\n新しいフォルダ名を入力してください（空白でルートフォルダ）:`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const folderName = response.getResponseText().trim();
  if (folderName) {
    try {
      const folders = DriveApp.getFoldersByName(folderName);
      const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
      PropertiesService.getScriptProperties().setProperty('PDF_FOLDER_ID', folder.getId());
      ui.alert('完了', `PDF出力フォルダを「${folderName}」に設定しました。`, ui.ButtonSet.OK);
    } catch (error) {
      ui.alert('エラー', 'フォルダの設定に失敗しました: ' + error.message, ui.ButtonSet.OK);
    }
  } else {
    PropertiesService.getScriptProperties().deleteProperty('PDF_FOLDER_ID');
    ui.alert('完了', 'PDF出力フォルダをルートフォルダに設定しました。', ui.ButtonSet.OK);
  }
}

function getPDFFolder() {
  try {
    const id = PropertiesService.getScriptProperties().getProperty('PDF_FOLDER_ID');
    if (id) return DriveApp.getFolderById(id);
  } catch (e) {
    console.error('フォルダ取得エラー:', e);
  }
  return DriveApp.getRootFolder();
}

// ===========================================
// オプション／マッピング
// ===========================================

function getCurrencySymbol(currency) {
  const symbols = {
    'USD': '$',
    'EUR': '€',
    'JPY': '¥',
    'GBP': '£',
    'CHF': 'CHF'
  };
  return symbols[currency] || '$';
}

function getDropdownOptions() {
  return {
    jewels: [
      '0 to 1 Jewels',
      '2 to 7 Jewels',
      '8 to 17 Jewels',
      'over 17 Jewels'
    ],
    bandMaterial: ['Textile', 'Metal', 'Leather', 'No Band'],
    caseMaterial: [
      'Gold/Silver Plated',
      'NOT Gold/Silver Plated',
      'Metal Clad w/Precious Metal',
      'Wholly of Precious Metal',
      'Other'
    ],
    backplateMaterial: ['Wholly of Precious Metal', 'Other'],
    countries: ['Japan', 'United states', 'Schweiz', 'Germany', 'China', 'Other'],
    primaryFunction: ['Timekeeping', 'GPS', 'Heart Monitor', 'Wi-Fi', 'Pedometer', 'Other']
  };
}

function getValueBreakoutConfig() {
  return {
    quartz: { movement: 0.50, case: 0.30, strap: 0.15, battery: 0.05 },
    mechanical: { movement: 0.55, case: 0.30, strap: 0.15, battery: 0.00 }
  };
}

function mapJewelsToDropdown(jewelCount) {
  if (typeof jewelCount === 'string') {
    const s = jewelCount.toLowerCase();
    if (s.includes('not applicable') || s.includes('quartz')) return '0 to 1 Jewels';
    const m = jewelCount.match(/\d+/);
    if (!m) return '0 to 1 Jewels';
    jewelCount = parseInt(m[0], 10);
  }
  if (jewelCount === 0 || jewelCount === 1) return '0 to 1 Jewels';
  if (jewelCount >= 2 && jewelCount <= 7) return '2 to 7 Jewels';
  if (jewelCount >= 8 && jewelCount <= 17) return '8 to 17 Jewels';
  if (jewelCount > 17) return 'over 17 Jewels';
  return '0 to 1 Jewels';
}

function mapCountryToDropdown(country) {
  if (!country) return '';
  const lower = String(country).toLowerCase();
  const opts = getDropdownOptions().countries;
  for (let o of opts) {
    if (lower.includes(o.toLowerCase())) return o;
  }
  if (lower.includes('switzerland') || lower.includes('swiss')) return 'Schweiz';
  if (lower.includes('usa') || lower.includes('america')) return 'United states';
  if (lower.includes('jp') || lower.includes('jpn')) return 'Japan';
  if (lower.includes('de') || lower.includes('deutsch')) return 'Germany';
  return 'Other';
}

// ===========================================
// ChatGPTデータ解析
// ===========================================

/** 外部UIから呼ぶ：AWB + ChatGPTテキストでシート作成 */
function createWatchWorksheet(awbNumber, chatgptData) {
  try {
    const parsed = parseChatGPTData(chatgptData);
    if (!parsed.success) return { success: false, message: parsed.message };
    parsed.data.awbNumber = awbNumber || '';
    const result = generateNewFormatWorksheet(parsed.data);
    return result;
  } catch (e) {
    console.error('ワークシート作成エラー:', e);
    return { success: false, message: 'ワークシート作成中にエラーが発生しました: ' + e.message };
  }
}

function parseChatGPTData(rawData) {
  try {
    if (!rawData || typeof rawData !== 'string') {
      return { success: false, message: 'データが無効です。ChatGPTからの出力をコピー・ペーストしてください。' };
    }
    const startMarker = '=== WATCH WORKSHEET DATA ===';
    const endMarker = '=== END DATA ===';
    const startIndex = rawData.indexOf(startMarker);
    const endIndex = rawData.indexOf(endMarker);
    if (startIndex === -1 || endIndex === -1) {
      return {
        success: false,
        message: 'データ形式が正しくありません。\n「=== WATCH WORKSHEET DATA ===」から「=== END DATA ===」までの部分が見つかりません。'
      };
    }
    const dataSection = rawData.substring(startIndex + startMarker.length, endIndex).trim();
    const lines = dataSection.split('\n').filter(l => l.trim());
    const parsed = {};
    for (let line of lines) {
      const idx = line.indexOf(':');
      if (idx === -1) continue;
      const key = line.substring(0, idx).trim();
      const value = line.substring(idx + 1).trim();
      if (value && value !== '[要確認]' && value !== '[不明]') parsed[key] = value;
    }

    const required = ['Style name/No/Reference', 'Total Watch Value', 'Movement Type'];
    const missing = required.filter(f => !parsed[f]);
    if (missing.length > 0) return { success: false, message: `必須項目が不足しています: ${missing.join(', ')}` };

    const normalized = normalizeData(parsed);
    const chk = validateData(normalized);
    if (!chk.isValid) return { success: false, message: '検証エラー: ' + chk.errors.join(' / ') };

    return { success: true, data: normalized };
  } catch (e) {
    console.error('データ解析エラー:', e);
    return { success: false, message: 'データの解析中にエラーが発生しました: ' + e.message };
  }
}

function normalizeData(raw) {
  const n = {};
  n.styleRef = raw['Style name/No/Reference'] || '';

  // 通貨単位を抽出して保存
  const totalValueStr = String(raw['Total Watch Value'] || '').trim();
  n.totalValue = parseFloat(totalValueStr.replace(/[^0-9.]/g, '')) || 0;

  // 通貨単位を抽出（JPY, USD, EUR等）
  const currencyMatch = totalValueStr.match(/[A-Z]{3}$/);
  n.currency = currencyMatch ? currencyMatch[0] : 'USD';

  n.movementType = raw['Movement Type'] || 'Quartz';
  n.htsCode = cleanHTSCode(raw['HTS US Code']) || '';
  n.jewels = mapJewelsToDropdown(raw['Number of Jewels in Movement']);
  n.quantity = parseInt(raw['Quantity'] || '1', 10);

  n.bandMaterial = raw['Material of Band'] || 'Leather';
  n.bandDetail = raw['Band Detail'] || '';
  n.caseMaterial = raw['Material of Case'] || 'Other';
  n.caseDetail = raw['Case Detail'] || '';
  n.backplateMaterial = raw['Material of Backplate'] || 'Other';
  n.backplateDetail = raw['Backplate Detail'] || '';

  n.movementCountry = mapCountryToDropdown(raw['Country of Origin of Movement']);
  n.bandCountry = mapCountryToDropdown(raw['Country of Origin of Band']);
  n.caseCountry = mapCountryToDropdown(raw['Country of Origin of Case']);
  n.batteryCountry = mapCountryToDropdown(raw['Country of Origin of Battery'] || 'Japan');

  n.primaryFunction = raw['Primary Function'] || 'Timekeeping';
  n.otherMaterials = raw['Other materials'] || '';

  const b = calculateValueBreakout(n.totalValue, n.movementType);
  n.movementValue = b.movement;
  n.caseValue = b.case;
  n.strapValue = b.strap;
  n.batteryValue = b.battery;

  const cfg = getCompanyConfig();
  n.companyName = cfg.companyName || '';
  n.nameAndTitle = cfg.nameAndTitle || '';
  n.email = cfg.email || '';
  n.awbNumber = '';

  return n;
}

function cleanHTSCode(htsText) {
  if (!htsText) return '';
  let cleaned = String(htsText).replace(/\s*\([^)]*\)/g, '');
  const m = cleaned.match(/\d+(\.\d+)*/);
  if (m) return m[0];
  return cleaned.trim().split(/\s+/)[0];
}

function calculateValueBreakout(totalValue, movementType) {
  const cfg = getValueBreakoutConfig();
  const r = (String(movementType).toLowerCase().includes('quartz')) ? cfg.quartz : cfg.mechanical;
  const round2 = (x) => Math.round(x * 100) / 100;
  return {
    movement: round2(totalValue * r.movement),
    case: round2(totalValue * r.case),
    strap: round2(totalValue * r.strap),
    battery: round2(totalValue * r.battery)
  };
}

function validateData(data) {
  const errors = [];
  if (!data.styleRef) errors.push('Style name/No/Reference は必須です');
  if (!data.totalValue || data.totalValue <= 0) errors.push('Total Watch Value は正の数値である必要があります');
  if (!data.movementType) errors.push('Movement Type は必須です');
  return { isValid: errors.length === 0, errors };
}

// ===========================================
// FedEx新形式ワークシート 生成（横型レイアウト）
// ===========================================

function generateNewFormatWorksheet(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
    const sheetName = 'NewWatch_' + ts;
    const sheet = ss.insertSheet(sheetName);

    createNewFormatTemplate(sheet);
    populateNewFormatData(sheet, data, 1);
    formatNewFormatWorksheet(sheet);

    ss.setActiveSheet(sheet);
    return { success: true, message: 'Watch Worksheet「' + sheetName + '」を作成しました。', sheetName };
  } catch (e) {
    console.error('ワークシート生成エラー:', e);
    return { success: false, message: 'ワークシートの生成中にエラーが発生しました: ' + e.message };
  }
}

function createNewFormatTemplate(sheet) {
  sheet.clear();

  // 2列のみ（A列：ラベル、B列：Watch 1データ）
  for (let i = 1; i <= 2; i++) {
    sheet.setColumnWidth(i, 450);
  }

  const titleRange = sheet.getRange(1, 1, 1, 2);
  titleRange.merge();
  titleRange.setValue('Watch Worksheet');
  titleRange.setFontSize(14);
  titleRange.setFontWeight('bold');
  titleRange.setHorizontalAlignment('center');
  titleRange.setBackground('#4472C4');
  titleRange.setFontColor('#FFFFFF');

  // Watch 1のみ
  const headers = ['', 'Watch 1'];
  for (let i = 0; i < headers.length; i++) {
    sheet.getRange(2, i + 1).setValue(headers[i]);
  }

  // 新形式：全て記述式（チェックボックスなし）
  const labels = [
    'Style name/No/Reference',
    'Style of watch',
    '  If Other, provide type',
    'Quantity',
    'HTSUS Number (if known)',
    'HTSUS Number (if known)',
    'HTSUS Number (if known)',
    'HTSUS Number (if known)',
    'What is the primary function of watch',
    '  If Other, provide primary function',
    'How is the watch powered',
    'Country of Origin of the battery',
    'Movement/ Display type',
    'Is the movement\'s size over 12mm in thickness and 50mm in width, length, or diameter?',
    'Number of Jewels in Movement',
    'Country of Origin of Movement',
    'Material of Band (Strap)',
    '  If Leather, provide type of animal',
    '  If Metal, provide type of metal',
    '  If Other, provide material',
    'Country of Origin of Band (Strap)',
    'Material of Case',
    '  If Other, provide material',
    'Country of Origin of Case',
    'Material of Backplate',
    '  If Other, provide material',
    'Value Breakout (amount and currency)',
    '  Movement',
    '  Case',
    '  Strap',
    '  Battery',
    '  Total Watch Value',
    '',
    'Company Name',
    'Name and Title',
    'E-mail',
    'AWB Number'
  ];

  let row = 3;
  for (let label of labels) {
    sheet.getRange(row, 1).setValue(label);
    row++;
  }
}

function populateNewFormatData(sheet, data, watchColumn) {
  const col = watchColumn + 1;

  // 行3: Style name/No/Reference
  sheet.getRange(3, col).setValue(data.styleRef || '');

  // 行4: Style of watch（記述式：Wrist / Pocket / Other）
  sheet.getRange(4, col).setValue('Wrist');

  // 行5: If Other, provide type（Wrist/Pocket以外の場合のみ）
  // 通常は空欄

  // 行6: Quantity
  sheet.getRange(6, col).setValue(data.quantity || 1);

  // 行7-10: HTSUS Number (if known) - 4行
  const htsNumeric = (data.htsCode || '').replace(/\./g, '');
  sheet.getRange(7, col).setValue(htsNumeric);

  // 行11: What is the primary function of watch（記述式）
  sheet.getRange(11, col).setValue(data.primaryFunction || 'Timekeeping');

  // 行12: If Other, provide primary function
  if (data.primaryFunction && !['Timekeeping', 'GPS', 'Heart Monitor', 'Wi-Fi', 'Pedometer'].includes(data.primaryFunction)) {
    sheet.getRange(12, col).setValue(data.primaryFunction);
    sheet.getRange(11, col).setValue('Other');
  }

  // 行13: How is the watch powered（記述式）
  const isQuartz = String(data.movementType).toLowerCase().includes('quartz');
  if (isQuartz) {
    sheet.getRange(13, col).setValue('Electric (Battery)');
  } else if (String(data.movementType).toLowerCase().includes('automatic')) {
    sheet.getRange(13, col).setValue('Automatic Winding (Self Winding)');
  } else {
    sheet.getRange(13, col).setValue('Manual');
  }

  // 行14: Country of Origin of the battery
  if (isQuartz) {
    sheet.getRange(14, col).setValue(data.batteryCountry || 'Japan');
  } else {
    sheet.getRange(14, col).setValue('N/A');
  }

  // 行15: Movement/ Display type（記述式）
  sheet.getRange(15, col).setValue(data.movementType || '');

  // 行16: Is the movement's size over 12mm...（記述式：Yes/No）
  sheet.getRange(16, col).setValue('No');

  // 行17: Number of Jewels in Movement（数値直接入力）
  let jewelCount = 0;
  if (data.jewels) {
    if (data.jewels === '0 to 1 Jewels') jewelCount = 0;
    else if (data.jewels === '2 to 7 Jewels') jewelCount = 5;
    else if (data.jewels === '8 to 17 Jewels') jewelCount = 12;
    else if (data.jewels === 'over 17 Jewels') jewelCount = 21;
  }
  sheet.getRange(17, col).setValue(jewelCount);

  // 行18: Country of Origin of Movement
  sheet.getRange(18, col).setValue(data.movementCountry || '');

  // 行19: Material of Band (Strap)（記述式）
  sheet.getRange(19, col).setValue(data.bandMaterial || '');

  // 行20-22: Band詳細（該当する行に記入）
  if (data.bandMaterial === 'Leather' && data.bandDetail) {
    sheet.getRange(20, col).setValue(data.bandDetail);
  }
  if (data.bandMaterial === 'Metal' && data.bandDetail) {
    sheet.getRange(21, col).setValue(data.bandDetail);
  }
  if (data.bandMaterial && !['Textile', 'Metal', 'Leather', 'No Band'].includes(data.bandMaterial)) {
    sheet.getRange(22, col).setValue(data.bandMaterial);
    sheet.getRange(19, col).setValue('Other');
  }

  // 行23: Country of Origin of Band (Strap)
  sheet.getRange(23, col).setValue(data.bandCountry || '');

  // 行24: Material of Case（記述式）
  sheet.getRange(24, col).setValue(data.caseMaterial || '');

  // 行25: If Other, provide material（Case）
  if (data.caseMaterial === 'Other' && data.caseDetail) {
    sheet.getRange(25, col).setValue(data.caseDetail);
  }

  // 行26: Country of Origin of Case
  sheet.getRange(26, col).setValue(data.caseCountry || '');

  // 行27: Material of Backplate（記述式）
  sheet.getRange(27, col).setValue(data.backplateMaterial || '');

  // 行28: If Other, provide material（Backplate）
  if (data.backplateMaterial !== 'Wholly of Precious Metal' && data.backplateDetail) {
    sheet.getRange(28, col).setValue(data.backplateDetail);
  }

  // 行29-34: Value Breakout
  const currency = data.currency || 'USD';
  // 行29はラベルのみ
  sheet.getRange(30, col).setValue(data.movementValue ? data.movementValue.toFixed(2) + ' ' + currency : '');
  sheet.getRange(31, col).setValue(data.caseValue ? data.caseValue.toFixed(2) + ' ' + currency : '');
  sheet.getRange(32, col).setValue(data.strapValue ? data.strapValue.toFixed(2) + ' ' + currency : '');
  sheet.getRange(33, col).setValue(data.batteryValue ? data.batteryValue.toFixed(2) + ' ' + currency : '');
  sheet.getRange(34, col).setValue(data.totalValue ? data.totalValue.toFixed(2) + ' ' + currency : '');

  // 行35は空行

  // 行36-39: Company info
  if (watchColumn === 1) {
    sheet.getRange(36, col).setValue(data.companyName || '');
    sheet.getRange(37, col).setValue(data.nameAndTitle || '');
    sheet.getRange(38, col).setValue(data.email || '');
    sheet.getRange(39, col).setValue(data.awbNumber || '');
  }
}

function formatNewFormatWorksheet(sheet) {
  const lastRow = 39;
  const lastCol = 2;  // 2列のみ（A列、B列）

  sheet.getRange(1, 1, lastRow, lastCol).setFontFamily('Arial').setFontSize(9);
  sheet.getRange(1, 1, 1, lastCol).setFontSize(14);
  sheet.setRowHeight(1, 30);

  const headerRange = sheet.getRange(2, 1, 1, lastCol);
  headerRange.setBackground('#4472C4')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(2, 30);

  const labelRange = sheet.getRange(3, 1, lastRow - 2, 1);
  labelRange.setFontWeight('bold')
    .setBackground('#E7E6E6')
    .setVerticalAlignment('middle');

  for (let row = 3; row <= lastRow; row++) {
    const label = sheet.getRange(row, 1).getValue();
    if (String(label).startsWith('  ')) {
      sheet.getRange(row, 1).setFontWeight('normal')
        .setBackground('#F2F2F2')
        .setFontStyle('italic');
    }
  }

  const dataRange = sheet.getRange(3, 2, lastRow - 2, lastCol - 1);
  dataRange.setBackground('#FFFFFF')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange(1, 1, lastRow, lastCol)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  for (let row = 3; row <= lastRow; row++) {
    sheet.setRowHeight(row, 20);
  }

  // Value Breakout セクション（行29-34）を黄色背景に
  sheet.getRange(29, 1, 6, lastCol).setBackground('#FFF2CC');
  // Company info セクション（行36-39）を青背景に
  sheet.getRange(36, 1, 4, lastCol).setBackground('#D9E1F2');
}

// ===========================================
// FedEx新形式PDF出力
// ===========================================

function exportCurrentSheetToNewFormatPDF() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const name = sheet.getName();
    if (!name.startsWith('NewWatch_')) {
      SpreadsheetApp.getUi().alert('エラー', '新形式Watch Worksheetのシートを選択してください。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    const result = exportToPDF(name, 'landscape');
    if (result.success) {
      SpreadsheetApp.getUi().alert('完了', `PDF「${result.fileName}」をGoogle Driveに保存しました。\nファイルID: ${result.fileId}`, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('エラー', result.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (e) {
    console.error('PDF出力エラー:', e);
    SpreadsheetApp.getUi().alert('エラー', 'PDF出力中にエラーが発生しました: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/** PDF出力（空白PDF対策の flush あり） */
function exportToPDF(sheetName, orientation) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error('指定されたシートが見つかりません: ' + sheetName);

    const original = ss.getActiveSheet();
    ss.setActiveSheet(sheet);
    SpreadsheetApp.flush();

    orientation = orientation || 'portrait';
    const isLandscape = orientation === 'landscape';
    const isDHL = sheetName.startsWith('DHL_') || sheetName === 'DHL用ver.2';
    const isDHLv2 = sheetName === 'DHL用ver.2';

    let url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?` +
      `exportFormat=pdf&format=pdf&` +
      (isDHL ? 'size=letter' : 'size=letter') +
      `&portrait=${!isLandscape}&fitw=true&`;

    if (isDHL) {
      url += isDHLv2
        ? 'top_margin=0.25&bottom_margin=0.25&left_margin=0.35&right_margin=0.35'
        : 'top_margin=0.35&bottom_margin=0.35&left_margin=0.6&right_margin=0.5';
    } else {
      url += 'top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5';
    }

    url += (isDHL && !isDHLv2 ? `&horizontal_alignment=LEFT` : `&horizontal_alignment=CENTER`) +
      `&vertical_alignment=TOP&gridlines=${isDHL ? 'false' : 'true'}&printtitle=false&sheetnames=false&` +
      `pagenum=UNDEFINED&attachment=false&gid=${sheet.getSheetId()}`;

    if (isDHL) {
      url += '&scale=4';
    }

    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });

    ss.setActiveSheet(original);

    const folder = getPDFFolder();
    const blob = response.getBlob().setName(`${sheetName}.pdf`);
    const file = folder.createFile(blob);

    return { success: true, fileId: file.getId(), fileName: file.getName(), fileUrl: file.getUrl() };
  } catch (e) {
    console.error('PDF エクスポートエラー:', e);
    return { success: false, message: 'PDF エクスポート中にエラーが発生しました: ' + e.message };
  }
}

// ===========================================
// DHL形式関連の関数
// ===========================================

/** 連続アンダースコアを作る（等幅フォント前提） */
function _us(n) { return new Array(Math.max(0, n)).fill('_').join(''); }

/** 下線の中に値を"載せる"テキストを作る（例: "__123.45____"） */
function _inlineOnLine(valueStr, totalLen, leftPad) {
  const v = String(valueStr || '').trim();
  const lp = Math.max(0, leftPad || 0);
  const usable = Math.max(0, totalLen - lp);
  if (!v) return _us(totalLen);
  const vv = v.length > usable ? v.substring(0, usable) : v;
  const right = Math.max(0, totalLen - lp - vv.length);
  return _us(lp) + vv + _us(right);
}

/** Name & Title の分離（簡易） */
function _splitNameAndTitle(nameAndTitle) {
  const s = String(nameAndTitle || '').trim();
  if (!s) return { name: '', title: '' };
  const parts = s.split(/\s+/);
  if (parts.length >= 3) {
    return { name: parts.slice(0, -1).join(' '), title: parts[parts.length - 1] };
  }
  if (parts.length === 2) {
    return { name: s, title: '' };
  }
  return { name: s, title: '' };
}

function exportCurrentSheetToDHLPDF() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const name = sheet.getName();
    if (!name.startsWith('NewWatch_')) {
      SpreadsheetApp.getUi().alert('エラー', '新形式Watch Worksheetのシートを選択してください。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    const data = extractDataFromNewFormatSheet(sheet);
    const dhlName = createDHLWorksheet(data);
    const result = exportToPDF(dhlName, 'portrait');

    const ui = SpreadsheetApp.getUi();
    const resp = ui.alert('確認', 'DHL形式のシートを残しますか？', ui.ButtonSet.YES_NO);
    if (resp === ui.Button.NO) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const dhl = ss.getSheetByName(dhlName);
      if (dhl) ss.deleteSheet(dhl);
    }

    if (result.success) {
      SpreadsheetApp.getUi().alert('完了', `DHL形式PDF「${result.fileName}」をGoogle Driveに保存しました。\nファイルID: ${result.fileId}`, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('エラー', result.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (e) {
    console.error('DHL PDF出力エラー:', e);
    SpreadsheetApp.getUi().alert('エラー', 'DHL形式PDF出力中にエラーが発生しました: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function exportCurrentSheetToDHLPDFv2() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const name = sheet.getName();
    if (!name.startsWith('NewWatch_')) {
      SpreadsheetApp.getUi().alert('エラー', '新形式Watch Worksheetのシートを選択してください。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    const data = extractDataFromNewFormatSheet(sheet);
    const v2Name = createDHLWorksheetV2(data);
    const result = exportToPDF(v2Name, 'portrait');

    const ui = SpreadsheetApp.getUi();
    const resp = ui.alert('確認', 'DHL用ver.2シートを残しますか？', ui.ButtonSet.YES_NO);
    if (resp === ui.Button.NO) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const v2 = ss.getSheetByName(v2Name);
      if (v2) ss.deleteSheet(v2);
    }

    if (result.success) {
      SpreadsheetApp.getUi().alert('完了', `DHL用ver.2 PDF「${result.fileName}」をGoogle Driveに保存しました。\nファイルID: ${result.fileId}`, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('エラー', result.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (e) {
    console.error('DHL ver.2 PDF出力エラー:', e);
    SpreadsheetApp.getUi().alert('エラー', 'DHL用ver.2のPDF出力中にエラーが発生しました: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function extractDataFromNewFormatSheet(sheet) {
  const data = {};

  // 行3: Style name/No/Reference
  data.styleRef = sheet.getRange(3, 2).getValue() || '';

  // 行6: Quantity
  data.quantity = sheet.getRange(6, 2).getValue() || 1;

  // 行7: HTSUS Number
  data.htsCode = sheet.getRange(7, 2).getValue() || '';

  // 行11: Primary function（記述式）
  data.primaryFunction = sheet.getRange(11, 2).getValue() || 'Timekeeping';

  // 行13: How is the watch powered（記述式）
  const powerSource = String(sheet.getRange(13, 2).getValue() || '').toLowerCase();
  if (powerSource.includes('electric') || powerSource.includes('battery')) {
    data.movementType = 'Quartz';
  } else if (powerSource.includes('automatic') || powerSource.includes('self')) {
    data.movementType = 'Automatic';
  } else {
    data.movementType = 'Mechanical';
  }

  // 行14: Country of Origin of the battery
  data.batteryCountry = sheet.getRange(14, 2).getValue() || 'Japan';

  // 行15: Movement/ Display type（記述式）
  const movementDisplay = sheet.getRange(15, 2).getValue() || '';
  if (movementDisplay) {
    data.movementType = movementDisplay;
  }

  // 行17: Number of Jewels in Movement（数値）
  const jewelCount = parseInt(sheet.getRange(17, 2).getValue() || '0', 10);
  if (jewelCount <= 1) data.jewels = '0 to 1 Jewels';
  else if (jewelCount <= 7) data.jewels = '2 to 7 Jewels';
  else if (jewelCount <= 17) data.jewels = '8 to 17 Jewels';
  else data.jewels = 'over 17 Jewels';

  // 行18: Country of Origin of Movement
  data.movementCountry = sheet.getRange(18, 2).getValue() || '';

  // 行19: Material of Band (Strap)（記述式）
  data.bandMaterial = sheet.getRange(19, 2).getValue() || '';

  // 行20-22: Band詳細
  data.bandDetail = sheet.getRange(20, 2).getValue() || sheet.getRange(21, 2).getValue() || sheet.getRange(22, 2).getValue() || '';

  // 行23: Country of Origin of Band (Strap)
  data.bandCountry = sheet.getRange(23, 2).getValue() || '';

  // 行24: Material of Case（記述式）
  data.caseMaterial = sheet.getRange(24, 2).getValue() || '';

  // 行25: Case詳細
  data.caseDetail = sheet.getRange(25, 2).getValue() || '';

  // 行26: Country of Origin of Case
  data.caseCountry = sheet.getRange(26, 2).getValue() || '';

  // 行27: Material of Backplate（記述式）
  data.backplateMaterial = sheet.getRange(27, 2).getValue() || '';

  // 行28: Backplate詳細
  data.backplateDetail = sheet.getRange(28, 2).getValue() || '';

  // 行30-34: Values - 通貨単位も抽出
  const totalValueStr = String(sheet.getRange(34, 2).getValue() || '').trim();
  const currencyMatch = totalValueStr.match(/[A-Z]{3}$/);
  data.currency = currencyMatch ? currencyMatch[0] : 'USD';

  data.movementValue = parseFloat(String(sheet.getRange(30, 2).getValue() || '').replace(/[^0-9.]/g, '')) || 0;
  data.caseValue = parseFloat(String(sheet.getRange(31, 2).getValue() || '').replace(/[^0-9.]/g, '')) || 0;
  data.strapValue = parseFloat(String(sheet.getRange(32, 2).getValue() || '').replace(/[^0-9.]/g, '')) || 0;
  data.batteryValue = parseFloat(String(sheet.getRange(33, 2).getValue() || '').replace(/[^0-9.]/g, '')) || 0;
  data.totalValue = parseFloat(String(sheet.getRange(34, 2).getValue() || '').replace(/[^0-9.]/g, '')) || 0;

  // 行36-39: Company info
  data.companyName = sheet.getRange(36, 2).getValue() || '';
  data.nameAndTitle = sheet.getRange(37, 2).getValue() || '';
  data.email = sheet.getRange(38, 2).getValue() || '';
  data.awbNumber = sheet.getRange(39, 2).getValue() || '';

  return data;
}

function createDHLWorksheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  const sheetName = 'DHL_Watch_' + ts;
  const sheet = ss.insertSheet(sheetName);

  createDHLTemplate(sheet);
  populateDHLData(sheet, data);
  formatDHLWorksheet(sheet);

  return sheetName;
}

function createDHLTemplate(sheet) {
  sheet.clear();
  sheet.setHiddenGridlines(true);

  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 460);

  sheet.getRange('A1:C80').setFontSize(10).setWrap(false);
  sheet.getRange('A1:A80').setFontFamily('Arial');
  sheet.getRange('B1:B80').setFontFamily('Arial')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange('C1:C80').setFontFamily('Courier New');

  sheet.getRange('A1:C1').merge()
    .setValue('WATCH / CLOCK WORKSHEET')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setFontWeight('bold').setFontSize(14);
  sheet.setRowHeight(1, 28);

  sheet.getRange('A3').setValue('From:');
  sheet.getRange('C3').setValue(_us(26) + '   ' + 'Date: ' + _us(30));
  sheet.setRowHeight(3, 20);

  sheet.getRange('A4').setValue('Airway bill#:');
  sheet.getRange('C4').setValue(_us(21));
  sheet.setRowHeight(4, 20);

  sheet.getRange('A6:C6').merge()
    .setValue('In order to clear this shipment with U.S. Customs, we need the following information for each type of watch / clock in this shipment.')
    .setWrap(true)
    .setVerticalAlignment('middle');
  sheet.setRowHeight(6, 36);

  sheet.getRange('A9').setValue('Movement:').setFontWeight('bold');
  sheet.getRange('A10').setValue('Type of movement:').setFontWeight('bold');

  const cbRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange('B11').setDataValidation(cbRule);
  sheet.getRange('C11').setValue('Analogue (Mechanical)');
  sheet.getRange('B12').setDataValidation(cbRule);
  sheet.getRange('C12').setValue('Digital (Electronic)');

  sheet.getRange('A13').setValue('Value:').setFontWeight('bold');
  sheet.getRange('C13').setValue(_inlineOnLine('', 16, 0));
  sheet.getRange('A14').setValue('Number of jewels:').setFontWeight('bold');
  sheet.getRange('C14').setValue(_inlineOnLine('', 21, 0));

  sheet.getRange('A16').setValue('Case:').setFontWeight('bold');
  sheet.getRange('A17').setValue('Material made of:').setFontWeight('bold');
  sheet.getRange('C17').setValue(_inlineOnLine('', 61, 0));
  sheet.getRange('A18').setValue('Value:').setFontWeight('bold');
  sheet.getRange('C18').setValue(_inlineOnLine('', 16, 0));

  sheet.getRange('A20').setValue('Strap:').setFontWeight('bold');
  sheet.getRange('A21').setValue('Material made of:').setFontWeight('bold');
  sheet.getRange('C21').setValue(_inlineOnLine('', 61, 0));
  sheet.getRange('A22').setValue('Value:').setFontWeight('bold');
  sheet.getRange('C22').setValue(_inlineOnLine('', 16, 0));

  sheet.getRange('A24').setValue('Type of Power:').setFontWeight('bold');
  sheet.getRange('B25').setDataValidation(cbRule); sheet.getRange('C25').setValue('Battery');
  sheet.getRange('B28').setDataValidation(cbRule); sheet.getRange('C28').setValue('Wind up');
  sheet.getRange('B29').setDataValidation(cbRule); sheet.getRange('C29').setValue('Self-Winding');

  sheet.getRange('A26').setValue('Type of Battery:').setFontWeight('bold');
  sheet.getRange('C26').setValue(_inlineOnLine('', 50, 0));
  sheet.getRange('A27').setValue('Value:').setFontWeight('bold');
  sheet.getRange('C27').setValue(_inlineOnLine('', 14, 0));

  sheet.getRange('A31').setValue('Signature:').setFontWeight('bold');
  sheet.getRange('C31').setValue(_inlineOnLine('', 74, 0));
  sheet.getRange('A33').setValue('Printed Name:').setFontWeight('bold');
  sheet.getRange('C33').setValue(_inlineOnLine('', 74, 0));
  sheet.getRange('A35').setValue('Company:').setFontWeight('bold');
  sheet.getRange('C35').setValue(_inlineOnLine('', 74, 0));

  sheet.getRange('A38').setValue('HTS information for entry: (For DHL use only)').setFontWeight('bold');
  sheet.getRange('A40').setValue('xxxx.xx.xx10 = Movement');
  sheet.getRange('A41').setValue('xxxx.xx.xx20 = Case');
  sheet.getRange('A42').setValue('xxxx.xx.xx30 = Strap');
  sheet.getRange('A43').setValue('xxxx.xx.xx40 = Power/Battery');
}

function populateDHLData(sheet, data) {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'MM/dd/yyyy');

  const fromLine = _inlineOnLine(data.companyName || '', 26, 0);
  const dateLine = _inlineOnLine(today, 30, 0);
  sheet.getRange('C3').setValue(fromLine + '   ' + 'Date: ' + dateLine);
  sheet.getRange('C4').setValue(_inlineOnLine(data.awbNumber || '', 21, 0));

  const isQuartz = String(data.movementType || '').toLowerCase().includes('quartz');
  sheet.getRange('B11').setValue(!isQuartz);
  sheet.getRange('B12').setValue(!!isQuartz);

  const currency = data.currency || 'USD';
  if (data.movementValue) {
    const val = Number(data.movementValue).toFixed(2) + ' ' + currency;
    sheet.getRange('C13').setValue(_inlineOnLine(val, 24, 0));
  }
  let jewelText = '';
  if (data.jewels && data.jewels !== '0 to 1 Jewels') {
    if (data.jewels === '0 to 1 Jewels') jewelText = '0-1';
    else if (data.jewels === '2 to 7 Jewels') jewelText = '2-7';
    else if (data.jewels === '8 to 17 Jewels') jewelText = '8-17';
    else if (data.jewels === 'over 17 Jewels') jewelText = '17+';
  }
  sheet.getRange('C14').setValue(_inlineOnLine(jewelText, 21, 0));

  sheet.getRange('C17').setValue(_inlineOnLine(data.caseMaterial || '', 61, 0));
  if (data.caseValue) sheet.getRange('C18').setValue(_inlineOnLine(Number(data.caseValue).toFixed(2) + ' ' + currency, 24, 0));

  sheet.getRange('C21').setValue(_inlineOnLine(data.bandMaterial || '', 61, 0));
  if (data.strapValue) sheet.getRange('C22').setValue(_inlineOnLine(Number(data.strapValue).toFixed(2) + ' ' + currency, 24, 0));

  const hasBattery = (Number(data.batteryValue) || 0) > 0 || isQuartz;
  sheet.getRange('B25').setValue(!!hasBattery);
  sheet.getRange('B28').setValue(!isQuartz);
  sheet.getRange('B29').setValue(!isQuartz);

  sheet.getRange('C26').setValue(_inlineOnLine(hasBattery ? 'Standard watch battery' : '', 50, 0));
  if (hasBattery && data.batteryValue) {
    sheet.getRange('C27').setValue(_inlineOnLine(Number(data.batteryValue).toFixed(2) + ' ' + currency, 24, 0));
  }

  sheet.getRange('C33').setValue(_inlineOnLine(data.nameAndTitle || '', 74, 0));
  sheet.getRange('C35').setValue(_inlineOnLine(data.companyName || '', 74, 0));
}

function formatDHLWorksheet(sheet) {
  const h = {
    1:28, 3:20, 4:20, 6:36,
    9:18, 10:18, 11:18, 12:18, 13:18, 14:18,
    16:18, 17:18, 18:18,
    20:18, 21:18, 22:18,
    24:18, 25:18, 26:18, 27:18, 28:18, 29:18,
    31:18, 33:18, 35:18,
    38:18, 40:18, 41:18, 42:18, 43:18
  };
  Object.keys(h).forEach(r => sheet.setRowHeight(Number(r), h[r]));
  sheet.getRange('A1:C80').setWrap(false);
  sheet.getRange('A6:C6').setWrap(true);
}

function createDHLWorksheetV2(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'DHL用ver.2';

  const existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);

  const sheet = ss.insertSheet(sheetName);

  createDHLV2Template(sheet);
  populateDHLV2Data(sheet, data);
  formatDHLV2Worksheet(sheet);

  return sheetName;
}

function createDHLV2Template(sheet) {
  sheet.clear();
  sheet.setHiddenGridlines(true);

  sheet.setColumnWidth(1, 20);
  sheet.setColumnWidth(2, 390);
  sheet.setColumnWidth(3, 20);
  sheet.setColumnWidth(4, 70);
  sheet.setColumnWidth(5, 20);
  sheet.setColumnWidth(6, 70);

  sheet.getRange('A1:F160').setFontFamily('Courier New').setFontSize(9).setWrap(false);
  sheet.getRange('A1:A160').setFontFamily('Arial');
  sheet.getRange('C1:C160').setFontFamily('Arial');
  sheet.getRange('E1:E160').setFontFamily('Arial');

  const cbRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();

  sheet.getRange('A1:F1').merge()
    .setValue('Watch Worksheet')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setFontWeight('bold').setFontSize(13);
  sheet.setRowHeight(1, 24);

  sheet.getRange('A3:F4').merge().setWrap(true);
  sheet.getRange('A3').setValue('1. Description of goods:\nPart number:').setVerticalAlignment('top');

  sheet.getRange('A5:F5').merge().setValue('2. Tariff number (if known):');

  sheet.getRange('A6:F6').merge();
  sheet.getRange('A6').setValue('The Harmonized Tariff Schedule of the United States is online: http://www.usitc.gov/tata/hts/bychapter/index.htm').setWrap(true);

  sheet.getRange('A8:F8').merge().setValue('3. Which best describes the article?');
  sheet.getRange('A9').setDataValidation(cbRule);
  sheet.getRange('B9').setValue('Wrist watch');
  sheet.getRange('A10').setDataValidation(cbRule);
  sheet.getRange('B10').setValue('Pocket or other watches not worn on the wrist');

  sheet.getRange('A12:F12').merge().setValue('4. If the case has precious metal*, what type is it?');
  sheet.getRange('A13').setDataValidation(cbRule);
  sheet.getRange('B13').setValue('The case is wholly made of, or clad with, precious metal');
  sheet.getRange('A14').setDataValidation(cbRule);
  sheet.getRange('B14').setValue('The case is plated or inlaid with precious metal (Is it gold or silver plated?)');
  sheet.getRange('C14').setDataValidation(cbRule); sheet.getRange('D14').setValue('Yes');
  sheet.getRange('E14').setDataValidation(cbRule); sheet.getRange('F14').setValue('No');
  sheet.getRange('A15:F15').merge().setValue('* The precious metals are gold, silver, platinum, iridium, osmium, palladium, rhodium and ruthenium.');

  sheet.getRange('A17:F17').merge().setValue('5. Is the back plate wholly made of, or clad with, precious metal?');
  sheet.getRange('C17').setDataValidation(cbRule); sheet.getRange('D17').setValue('Yes');
  sheet.getRange('E17').setDataValidation(cbRule); sheet.getRange('F17').setValue('No');

  sheet.getRange('A19:F19').merge().setValue('6. What is the display?');
  sheet.getRange('A20').setDataValidation(cbRule); sheet.getRange('B20').setValue('Mechanical (analog) only');
  sheet.getRange('A21').setDataValidation(cbRule); sheet.getRange('B21').setValue('Opto-electronic (digital) only');
  sheet.getRange('A22').setDataValidation(cbRule); sheet.getRange('B22').setValue('Both analog and digital');

  sheet.getRange('A24:F24').merge().setValue('7. What is the power source?');
  sheet.getRange('A25').setDataValidation(cbRule); sheet.getRange('B25').setValue('Electric (battery, solar)');
  sheet.getRange('A26').setDataValidation(cbRule); sheet.getRange('B26').setValue('Automatic (self) winding');
  sheet.getRange('A27').setDataValidation(cbRule); sheet.getRange('B27').setValue('Manual winding');

  sheet.getRange('A29:F29').merge().setValue('8. What is the strap, band or bracelet (or chain for pocket watches) made of?');
  sheet.getRange('A30').setDataValidation(cbRule); sheet.getRange('B30').setValue('Precious metal or base metal clad with precious metal');
  sheet.getRange('A31').setDataValidation(cbRule); sheet.getRange('B31').setValue('Base metal (stainless steel, brass, etc.), whether or not plated with precious metal');
  sheet.getRange('A32').setDataValidation(cbRule); sheet.getRange('B32').setValue('Textile/Cloth');
  sheet.getRange('A33').setDataValidation(cbRule); sheet.getRange('B33').setValue('Leather – what type (bovine, equine, etc.)?');
  sheet.getRange('A34').setDataValidation(cbRule); sheet.getRange('B34').setValue('Other:');
  sheet.getRange('A35').setDataValidation(cbRule); sheet.getRange('B35').setValue('No strap, band or bracelet');

  sheet.getRange('A37:F37').merge().setValue('9. How many jewels are in the movement?');
  sheet.getRange('A38:F38').merge();
  sheet.getRange('A39:F39').merge().setValue('10. What is the country of origin of the movement that controls hours/minutes?');
  sheet.getRange('A40:F40').merge();
  sheet.getRange('A41:F41').merge().setValue('11. If the movement value is not over $15 each, does it measure over 15.2 mm?');
  sheet.getRange('C42').setDataValidation(cbRule); sheet.getRange('D42').setValue('Yes');
  sheet.getRange('E42').setDataValidation(cbRule); sheet.getRange('F42').setValue('No');

  sheet.getRange('A44:F44').merge().setValue('12. List functions in order of importance (time keeping, GPS, Wi-Fi, heart monitor, pedometer, etc.)?');
  sheet.getRange('A45:F46').merge().setWrap(true);

  sheet.getRange('A48:F48').merge().setValue('Value breakdown of components:');
  sheet.getRange('A49:F49').merge().setValue('Movement: $            Strap: $');
  sheet.getRange('A50:F50').merge().setValue('Case: $');
  sheet.getRange('A51:F51').merge().setValue('Battery: $');
  sheet.getRange('A52:F52').merge().setValue('Total watch value: $');

  sheet.getRange('A54:F54').merge().setValue('Completed by:');
  sheet.getRange('A55:F55').merge().setValue(
    'Name: ' + _inlineOnLine('', 30, 0) + '  ' +
    'Signature: ' + _inlineOnLine('', 34, 0) + '  ' +
    'Date: ' + _inlineOnLine('', 12, 0)
  );
  sheet.getRange('A56:F56').merge().setValue(
    'Title: ' + _inlineOnLine('', 28, 0) + '  ' +
    'Company: ' + _inlineOnLine('', 40, 0)
  );
  sheet.getRange('A55:F56').setFontFamily('Courier New');

  sheet.getRange('A58:F59').merge()
    .setValue('This form is not required by U.S. Customs & Border Protection.\nHowever, a detailed description of merchandise is required per 19CFR 141.86.')
    .setWrap(true);
  sheet.getRange('A60').setValue('May 18, 2015');
}

function populateDHLV2Data(sheet, data) {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'MM/dd/yyyy');

  const desc = (data.styleRef || '').trim();
  const part = '';
  const descText = '1. Description of goods:\n' + (desc || '') + '\nPart number: ' + (part || '');
  sheet.getRange('A3').setValue(descText);
  sheet.getRange('A3').setWrap(true);

  sheet.getRange('A5').setValue('2. Tariff number (if known): ' + (data.htsCode || ''));

  sheet.getRange('A9').setValue(true);
  sheet.getRange('A10').setValue(false);

  const cm = (data.caseMaterial || '').toUpperCase();
  const isWholly = cm.includes('WHOLLY') || cm.includes('CLAD');
  const isPlated = cm.includes('PLATED') || cm.includes('INLAID');
  const goldOrSilver = cm.includes('GOLD') || cm.includes('SILVER');
  sheet.getRange('A13').setValue(!!isWholly);
  sheet.getRange('A14').setValue(!!isPlated);
  sheet.getRange('C14').setValue(isPlated ? !!goldOrSilver : false);
  sheet.getRange('E14').setValue(isPlated ? !goldOrSilver : false);

  const back = (data.backplateMaterial || '').toLowerCase();
  const backPrecious = back.includes('gold') || back.includes('silver');
  sheet.getRange('C17').setValue(!!backPrecious);
  sheet.getRange('E17').setValue(!backPrecious);

  const isQuartz = String(data.movementType || '').toLowerCase().includes('quartz');
  sheet.getRange('A20').setValue(!isQuartz);
  sheet.getRange('A21').setValue(isQuartz);
  sheet.getRange('A22').setValue(false);

  const hasBattery = (Number(data.batteryValue) || 0) > 0 || isQuartz;
  sheet.getRange('A25').setValue(!!hasBattery);
  sheet.getRange('A26').setValue(!isQuartz);
  sheet.getRange('A27').setValue(false);

  const bm = (data.bandMaterial || '').toLowerCase();
  const isNoBand = bm.includes('no band') || bm.includes('no strap');
  const isLeather = bm.includes('leather');
  const isTextile = bm.includes('textile') || bm.includes('fabric');
  const isMetal = bm.includes('metal') || bm.includes('steel');
  sheet.getRange('A30').setValue(false);
  sheet.getRange('A31').setValue(!!isMetal);
  sheet.getRange('A32').setValue(!!isTextile);
  sheet.getRange('A33').setValue(!!isLeather);
  sheet.getRange('A34').setValue(!isMetal && !isTextile && !isLeather && !isNoBand && !!bm);
  sheet.getRange('A35').setValue(!!isNoBand);

  let jewelText = '';
  if (data.jewels && data.jewels !== '0 to 1 Jewels') {
    if (data.jewels === '0 to 1 Jewels') jewelText = '0-1';
    else if (data.jewels === '2 to 7 Jewels') jewelText = '2-7';
    else if (data.jewels === '8 to 17 Jewels') jewelText = '8-17';
    else if (data.jewels === 'over 17 Jewels') jewelText = '17+';
  }
  sheet.getRange('A38').setValue(jewelText);
  sheet.getRange('A40').setValue(data.movementCountry || '');
  sheet.getRange('C42').setValue(false);
  sheet.getRange('E42').setValue(false);

  const n2 = (v)=> (v!=null && v!=='' ? Number(v).toFixed(2) : '');
  const currency = data.currency || 'USD';
  const currencySymbol = getCurrencySymbol(currency);
  sheet.getRange('A49').setValue('Movement: ' + currencySymbol + (n2(data.movementValue)||'') + ' ' + currency + '            Strap: ' + currencySymbol + (n2(data.strapValue)||'') + ' ' + currency);
  sheet.getRange('A50').setValue('Case: ' + currencySymbol + (n2(data.caseValue)||'') + ' ' + currency);
  sheet.getRange('A51').setValue('Battery: ' + currencySymbol + (n2(data.batteryValue)||'') + ' ' + currency);
  sheet.getRange('A52').setValue('Total watch value: ' + currencySymbol + (n2(data.totalValue)||'') + ' ' + currency);

  const split = _splitNameAndTitle(data.nameAndTitle || '');
  sheet.getRange('A55').setValue(
    'Name: ' + _inlineOnLine(split.name, 30, 0) + '  ' +
    'Signature: ' + _inlineOnLine('', 34, 0) + '  ' +
    'Date: ' + _inlineOnLine(today, 12, 0)
  );
  sheet.getRange('A56').setValue(
    'Title: ' + _inlineOnLine(split.title, 28, 0) + '  ' +
    'Company: ' + _inlineOnLine(data.companyName || '', 40, 0)
  );
}

function formatDHLV2Worksheet(sheet) {
  const heights = {
    1:24,
    3:44,4:24,5:18,6:24,
    8:18,9:16,10:16,
    12:18,13:16,14:16,15:18,
    17:18,
    19:18,20:16,21:16,22:16,
    24:18,25:16,26:16,27:16,
    29:18,30:16,31:16,32:16,33:16,34:16,35:16,
    37:18,38:16,39:18,40:16,41:18,42:16,
    44:18,45:28,
    48:18,49:16,50:16,51:16,52:16,
    54:18,55:20,56:20,
    58:28,60:16
  };
  Object.keys(heights).forEach(r => sheet.setRowHeight(Number(r), heights[r]));
  sheet.getRange('A1:F160').setWrap(false);
  sheet.getRange('A3:F4').setWrap(true);
  sheet.getRange('A6:F6').setWrap(true);
  sheet.getRange('A45:F46').setWrap(true);
  sheet.getRange('A58:F59').setWrap(true);
}

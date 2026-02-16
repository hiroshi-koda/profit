/**
 * Sales Pulse - Server Side Code (v7.2: Fix Date Serialization & Sort Crash)
 */

const APP_NAME = "Sales Pulse";
const SHEET_DEALS = "Deals";
const SHEET_ACTIONS = "Actions";
const SHEET_SETTINGS = "Settings";
const SHEET_TARGETS = "Targets";
const SHEET_DEAL_LOGS = "DealLogs";
const SHEET_CONTACTS = "Contacts";
const SHEET_CONTENTS = "Contents";
const SHEET_BILLING_MASTER = "BillingMaster";
const SHEET_MONTHLY_TASKS = "MonthlyTasks";

const BILLING_MASTER_HEADERS = [
  'id',
  'dealName',
  'clientName',
  'genre',
  'contractAmount',
  'contractStartMonth',
  'contractEndMonth',
  'billingType',
  'pjNumber',
  'transactionType',
  'creditCheck',
  'boxUrl',
  'saifuStatus',
  'deliveryStatus',
  'invoiceStatus',
  'salesFlow',
  'allianceDataUsage',
  'paymentPartner',
  'paymentPattern',
  'updatedAt'
];
const MONTHLY_TASK_HEADERS = [
  'id',
  'targetMonth',
  'dealId',
  'taskType',
  'status',
  'saifuStatus',
  'notes',
  'boxUrl',
  'updatedAt'
];

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const sheetMap = {};
  sheets.forEach(s => sheetMap[s.getName()] = s);

  // Schema guard: ensure sheets have the latest columns (e.g., contactId / dealId linkage)
  if (sheetMap[SHEET_DEALS]) ensureDealsSchema(sheetMap[SHEET_DEALS]);
  if (sheetMap[SHEET_ACTIONS]) ensureActionsSchema(sheetMap[SHEET_ACTIONS]);
  if (sheetMap[SHEET_CONTACTS]) ensureContactsSchema(sheetMap[SHEET_CONTACTS]);
  if (sheetMap[SHEET_CONTENTS]) ensureContentsSchema(sheetMap[SHEET_CONTENTS]);
  if (!sheetMap[SHEET_BILLING_MASTER]) {
    sheetMap[SHEET_BILLING_MASTER] = ensureSheet(ss, SHEET_BILLING_MASTER, BILLING_MASTER_HEADERS);
  }
  if (!sheetMap[SHEET_MONTHLY_TASKS]) {
    sheetMap[SHEET_MONTHLY_TASKS] = ensureSheet(ss, SHEET_MONTHLY_TASKS, MONTHLY_TASK_HEADERS);
  }
  ensureBillingMasterSchema(sheetMap[SHEET_BILLING_MASTER]);
  ensureMonthlyTasksSchema(sheetMap[SHEET_MONTHLY_TASKS]);
  ensureBillingMasterRowsFromDeals(ss);

  // Helper to get data safely
  const getData = (name) => {
    const sheet = sheetMap[name];
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    // Get all values at once (headers + data)
    // Using getDataRange is faster than separate getRange calls
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const dataRows = values.slice(1);
    
    return dataRows.map(row => {
      let obj = {};
      headers.forEach((header, i) => {
        if (row[i] instanceof Date) obj[header] = formatDateStr(row[i]);
        else obj[header] = row[i];
      });
      return obj;
    });
  };

  // Skip ensureSheet for read performance. Sheets should exist in normal operation.
  // If they don't, we just return empty arrays. Initialization happens on write or manual setup.

  const schemaStatus = {
    deals: getMissingColumns(sheetMap[SHEET_DEALS], [
      'id', 'clientName', 'dealName', 'amount', 'occurrenceDate', 'quoteDate', 'status', 'engagementStatus',
      'genre', 'service', 'channel', 'billingType', 'billingDate', 'revenueMonth', 'billingStartMonth',
      'billingEndMonth', 'billingAmount', 'billingTotal', 'contactId', 'contactIds', 'updatedAt'
    ])
  };

  return {
    deals: getData(SHEET_DEALS),
    actions: getData(SHEET_ACTIONS),
    settings: getData(SHEET_SETTINGS),
    targets: getData(SHEET_TARGETS),
    dealLogs: getData(SHEET_DEAL_LOGS),
    contacts: getData(SHEET_CONTACTS),
    contents: getData(SHEET_CONTENTS),
    billingMaster: getData(SHEET_BILLING_MASTER),
    monthlyTasks: getData(SHEET_MONTHLY_TASKS),
    schemaStatus
  };
}

function getMissingColumns(sheet, expected) {
  if (!sheet) return { missing: expected.slice(), hasAll: false };
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return { missing: expected.slice(), hasAll: false };
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const missing = expected.filter(col => !headers.includes(col));
  return { missing, hasAll: missing.length === 0 };
}

// --- Data Setup Helpers ---

function setupDemoData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headers = {
    settings: ['id', 'type', 'category', 'name', 'updatedAt'],
    deals: ['id', 'clientName', 'dealName', 'amount', 'occurrenceDate', 'quoteDate', 'status', 'engagementStatus', 'genre', 'service', 'channel', 'billingType', 'billingDate', 'revenueMonth', 'billingStartMonth', 'billingEndMonth', 'billingAmount', 'billingTotal', 'contactId', 'contactIds', 'updatedAt'],
    actions: ['id', 'targetDate', 'channel', 'actionContent', 'isDone', 'genre', 'dealId', 'actionType', 'updatedAt'],
    targets: ['id', 'year', 'month', 'type', 'category', 'metric', 'value', 'updatedAt'],
    dealLogs: ['id', 'dealId', 'logDate', 'logType', 'content', 'updatedAt'],
    contacts: ['id', 'company', 'name', 'department', 'email', 'mobilePhone', 'updatedAt'],
    contents: ['id', 'title', 'category', 'summary', 'body', 'campaignPeriod', 'cta', 'updatedAt']
  };
  
  resetSheet(ss, SHEET_SETTINGS, headers.settings);
  initDefaultSettings(ss.getSheetByName(SHEET_SETTINGS));

  resetSheet(ss, SHEET_TARGETS, headers.targets);
  initDefaultTargets(ss.getSheetByName(SHEET_TARGETS));

  resetSheet(ss, SHEET_CONTACTS, headers.contacts);
  const dummyContacts = generateDummyContacts(10);
  ss.getSheetByName(SHEET_CONTACTS).getRange(2, 1, dummyContacts.length, dummyContacts[0].length).setValues(dummyContacts);

  resetSheet(ss, SHEET_CONTENTS, headers.contents);
  const dummyContents = generateDummyContents();
  ss.getSheetByName(SHEET_CONTENTS).getRange(2, 1, dummyContents.length, dummyContents[0].length).setValues(dummyContents);

  resetSheet(ss, SHEET_DEALS, headers.deals);
  const dummyDeals = generateDummyDeals(60);
  // Map Deals to random Contacts
  dummyDeals.forEach(deal => {
    // Inject contactId/contactIds before updatedAt
    // Original generateDummyDeals returns: [id, client, name, amount, occDate, quoteDate, status, engagementStatus, genre, service, channel, billingType, billingDate, revenueMonth, billingStartMonth, billingEndMonth, billingAmount, billingTotal, updatedAt]
    const updatedAt = deal.pop();
    const contact = dummyContacts[Math.floor(Math.random() * dummyContacts.length)];
    // Simple matching by company name if possible, otherwise random
    const matchedContact = dummyContacts.find(c => c[1] === deal[1]); // c[1] is company
    const primaryId = matchedContact ? matchedContact[0] : (Math.random() > 0.7 ? contact[0] : '');
    const contactIdsArr = primaryId ? [primaryId] : [];
    // 30% chance to add a second contact for demo purposes
    if (Math.random() > 0.7) {
      const extra = dummyContacts[Math.floor(Math.random() * dummyContacts.length)][0];
      if (extra && extra !== primaryId) contactIdsArr.push(extra);
    }
    deal.push(primaryId); // legacy single contact
    deal.push(contactIdsArr.join(','));
    deal.push(updatedAt);
  });
  ss.getSheetByName(SHEET_DEALS).getRange(2, 1, dummyDeals.length, dummyDeals[0].length).setValues(dummyDeals);

  resetSheet(ss, SHEET_ACTIONS, headers.actions);
  const dummyActions = generateDummyActions(30);
  ss.getSheetByName(SHEET_ACTIONS).getRange(2, 1, dummyActions.length, dummyActions[0].length).setValues(dummyActions);

  resetSheet(ss, SHEET_DEAL_LOGS, headers.dealLogs);
  const dealIds = dummyDeals.map(d => d[0]);
  const dummyLogs = generateDummyDealLogs(dealIds);
  ss.getSheetByName(SHEET_DEAL_LOGS).getRange(2, 1, dummyLogs.length, dummyLogs[0].length).setValues(dummyLogs);
}

function setupBillingSampleData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headers = {
    settings: ['id', 'type', 'category', 'name', 'updatedAt'],
    deals: ['id', 'clientName', 'dealName', 'amount', 'occurrenceDate', 'quoteDate', 'status', 'engagementStatus', 'genre', 'service', 'channel', 'billingType', 'billingDate', 'revenueMonth', 'billingStartMonth', 'billingEndMonth', 'billingAmount', 'billingTotal', 'contactId', 'contactIds', 'updatedAt'],
    actions: ['id', 'targetDate', 'channel', 'actionContent', 'isDone', 'genre', 'dealId', 'actionType', 'updatedAt'],
    targets: ['id', 'year', 'month', 'type', 'category', 'metric', 'value', 'updatedAt'],
    dealLogs: ['id', 'dealId', 'logDate', 'logType', 'content', 'updatedAt'],
    contacts: ['id', 'company', 'name', 'department', 'email', 'mobilePhone', 'updatedAt'],
    contents: ['id', 'title', 'category', 'summary', 'body', 'campaignPeriod', 'cta', 'updatedAt']
  };

  ensureSheet(ss, SHEET_SETTINGS, headers.settings);
  ensureSheet(ss, SHEET_DEALS, headers.deals);
  ensureSheet(ss, SHEET_ACTIONS, headers.actions);
  ensureSheet(ss, SHEET_TARGETS, headers.targets);
  ensureSheet(ss, SHEET_DEAL_LOGS, headers.dealLogs);
  ensureSheet(ss, SHEET_CONTACTS, headers.contacts);
  ensureSheet(ss, SHEET_CONTENTS, headers.contents);
  ensureSheet(ss, SHEET_BILLING_MASTER, BILLING_MASTER_HEADERS);
  ensureSheet(ss, SHEET_MONTHLY_TASKS, MONTHLY_TASK_HEADERS);

  ensureDealsSchema(ss.getSheetByName(SHEET_DEALS));
  ensureActionsSchema(ss.getSheetByName(SHEET_ACTIONS));
  ensureContactsSchema(ss.getSheetByName(SHEET_CONTACTS));
  ensureContentsSchema(ss.getSheetByName(SHEET_CONTENTS));
  ensureBillingMasterSchema(ss.getSheetByName(SHEET_BILLING_MASTER));
  ensureMonthlyTasksSchema(ss.getSheetByName(SHEET_MONTHLY_TASKS));

  const sampleSettings = [
    { id: 'setting-demo-001', type: 'Genre', category: '', name: '製薬：治験' },
    { id: 'setting-demo-002', type: 'Service', category: '運輸', name: 'well-harmoレポート' },
    { id: 'setting-demo-003', type: 'Channel', category: '製薬：マーケティング', name: '広告代理店連携' }
  ];
  sampleSettings.forEach(saveSetting);

  const sampleContacts = [
    { id: 'contact-demo-001', company: '株式会社テスト商事', name: '山田 太郎', department: '営業企画', email: 'yamada@example.com', mobilePhone: '090-1111-2222' },
    { id: 'contact-demo-002', company: 'メディリンク株式会社', name: '佐藤 花子', department: 'マーケティング', email: 'sato@example.com', mobilePhone: '090-3333-4444' },
    { id: 'contact-demo-003', company: '株式会社フューチャー製薬', name: '鈴木 健', department: '事業開発', email: 'suzuki@example.com', mobilePhone: '090-5555-6666' }
  ];
  sampleContacts.forEach(saveContact);

  const sampleDeals = [
    { id: 'billing-demo-001', clientName: '株式会社テスト商事', dealName: '物流分析ダッシュボード導入', amount: 1800000, occurrenceDate: '2026-01-10', quoteDate: '2026-01-15', status: 'contract', engagementStatus: 'contacting', genre: '運輸', service: 'well-harmoレポート', channel: '既存顧客', billingType: 'single', billingDate: '2026-02-20', revenueMonth: '2026-02', billingStartMonth: '', billingEndMonth: '', billingAmount: 1800000, billingTotal: 1800000, contactId: 'contact-demo-001', contactIds: 'contact-demo-001' },
    { id: 'billing-demo-002', clientName: 'メディリンク株式会社', dealName: '治験対象者募集キャンペーン運用', amount: 3600000, occurrenceDate: '2026-01-05', quoteDate: '2026-01-12', status: 'contract', engagementStatus: 'contacting', genre: '製薬：治験', service: 'リクルーティングサービス', channel: 'CMIC', billingType: 'monthly', billingDate: '2026-01-31', revenueMonth: '2026-01', billingStartMonth: '2026-01', billingEndMonth: '2026-03', billingAmount: 1200000, billingTotal: 3600000, contactId: 'contact-demo-002', contactIds: 'contact-demo-002' },
    { id: 'billing-demo-003', clientName: '株式会社フューチャー製薬', dealName: 'MR向け商談支援コンテンツ制作', amount: 2400000, occurrenceDate: '2026-02-01', quoteDate: '2026-02-05', status: 'proposal', engagementStatus: 'contacting', genre: '製薬：マーケティング', service: 'SmartPDCA', channel: '広告代理店連携', billingType: 'quarterly', billingDate: '2026-03-31', revenueMonth: '2026-03', billingStartMonth: '2026-03', billingEndMonth: '2026-09', billingAmount: 800000, billingTotal: 2400000, contactId: 'contact-demo-003', contactIds: 'contact-demo-003' }
  ];
  sampleDeals.forEach(saveDeal);

  const sampleActions = [
    { id: 'action-demo-001', targetDate: '2026-02-03', channel: '既存顧客', actionContent: '請求内容の最終確認', isDone: false, genre: '運輸', dealId: 'billing-demo-001', actionType: 'deal' },
    { id: 'action-demo-002', targetDate: '2026-02-05', channel: 'CMIC', actionContent: '月次請求の送付', isDone: true, genre: '製薬：治験', dealId: 'billing-demo-002', actionType: 'deal' },
    { id: 'action-demo-003', targetDate: '2026-02-08', channel: '広告代理店連携', actionContent: '四半期請求の金額調整', isDone: false, genre: '製薬：マーケティング', dealId: 'billing-demo-003', actionType: 'deal' }
  ];
  sampleActions.forEach(saveAction);

  const sampleTargets = [
    { id: 'target-demo-001', year: 2026, month: 2, type: 'Main', category: 'Total', metric: 'Sales', value: 12000000 },
    { id: 'target-demo-002', year: 2026, month: 2, type: 'Genre', category: '運輸', metric: 'Sales', value: 4500000 },
    { id: 'target-demo-003', year: 2026, month: 2, type: 'Genre', category: '製薬：治験', metric: 'Sales', value: 5500000 }
  ];
  sampleTargets.forEach(saveTarget);

  const sampleLogs = [
    { id: 'log-demo-001', dealId: 'billing-demo-001', logDate: '2026-02-01', logType: 'Meeting', content: '納品物確認ミーティングを実施。' },
    { id: 'log-demo-002', dealId: 'billing-demo-002', logDate: '2026-02-02', logType: 'Result', content: '請求先部署の確認が完了。' },
    { id: 'log-demo-003', dealId: 'billing-demo-003', logDate: '2026-02-03', logType: 'NextAction', content: '次回打ち合わせで請求条件を確定予定。' }
  ];
  sampleLogs.forEach(saveDealLog);

  const sampleContents = [
    { id: 'content-demo-001', title: '請求運用フローのご案内', category: '運用資料', summary: '請求処理と入金確認の流れ', body: '毎月末に請求書を発行し、翌月10日までに入金確認を行う運用です。', campaignPeriod: '通年', cta: '確認依頼' },
    { id: 'content-demo-002', title: '製薬向け請求テンプレート', category: 'テンプレート', summary: '製薬案件の請求書フォーマット', body: '製薬案件向けに必要な項目を含んだ請求テンプレートです。', campaignPeriod: '通年', cta: 'テンプレート送付' },
    { id: 'content-demo-003', title: 'パートナー手数料支払ルール', category: '社内ナレッジ', summary: '半期手数料支払いの条件', body: '3月・9月に手数料支払対象案件を集計し支払処理を実施します。', campaignPeriod: '通年', cta: 'ルール確認' }
  ];
  sampleContents.forEach(saveContent);

  ensureBillingMasterRowsFromDeals(ss);

  const sampleMasters = [
    { id: 'billing-demo-001', pjNumber: 'PJ-LOG-1001', transactionType: '新規', creditCheck: '済', boxUrl: 'https://example.com/box/billing-demo-001', saifuStatus: '済', deliveryStatus: '納品済', invoiceStatus: '請求済', salesFlow: 'パートナー', allianceDataUsage: '無', paymentPartner: '東都パートナーズ', paymentPattern: '3月・9月支払い' },
    { id: 'billing-demo-002', pjNumber: 'PJ-CT-2026-002', transactionType: '既存', creditCheck: '－', boxUrl: 'https://example.com/box/billing-demo-002', saifuStatus: '未', deliveryStatus: '対応中', invoiceStatus: '未', salesFlow: '', allianceDataUsage: '有', paymentPartner: '', paymentPattern: '' },
    { id: 'billing-demo-003', pjNumber: '', transactionType: '新規', creditCheck: '未', boxUrl: '', saifuStatus: '未', deliveryStatus: '未', invoiceStatus: '未', salesFlow: '', allianceDataUsage: '無', paymentPartner: '', paymentPattern: '' }
  ];
  sampleMasters.forEach(saveBillingMaster);

  const sampleTasks = [
    { id: 'billing-task-demo-001', targetMonth: '2026-02', dealId: 'billing-demo-001', taskType: '売上請求', status: '対応中', saifuStatus: '済', notes: '請求書送付済み。入金確認待ち。', boxUrl: 'https://example.com/box/task-001' },
    { id: 'billing-task-demo-002', targetMonth: '2026-02', dealId: 'billing-demo-002', taskType: '支払明細', status: '未', saifuStatus: '未', notes: '前月分のデータ連携明細を作成予定。', boxUrl: 'https://example.com/box/task-002' },
    { id: 'billing-task-demo-003', targetMonth: '2026-03', dealId: 'billing-demo-001', taskType: '手数料支払', status: '未', saifuStatus: '未', notes: '半期支払い予定。契約条件を再確認。', boxUrl: '' }
  ];
  sampleTasks.forEach(saveMonthlyTask);

  return {
    message: '全シートへサンプルデータ（各3件）を投入しました。',
    settings: sampleSettings.length,
    deals: sampleDeals.length,
    actions: sampleActions.length,
    targets: sampleTargets.length,
    dealLogs: sampleLogs.length,
    contacts: sampleContacts.length,
    contents: sampleContents.length,
    billingMasters: sampleMasters.length,
    monthlyTasks: sampleTasks.length
  };
}

function generateDummyContacts(count) {
  const companies = ['株式会社アルファ', 'ベータ商事', 'ガンマ製薬', 'デルタ物流', 'オメガ・マーケティング'];
  const firstNames = ['太郎', '次郎', '花子', '健太', '美咲'];
  const lastNames = ['佐藤', '鈴木', '高橋', '田中', '伊藤'];
  const rows = [];
  const now = new Date();

  for (let i = 0; i < count; i++) {
    const com = companies[Math.floor(Math.random() * companies.length)];
    const name = lastNames[Math.floor(Math.random() * lastNames.length)] + ' ' + firstNames[Math.floor(Math.random() * firstNames.length)];
    rows.push([
      Utilities.getUuid(),
      com,
      name,
      '営業部',
      `user${i+1}@example.com`,
      `090-${Math.floor(1000+Math.random()*9000)}-${Math.floor(1000+Math.random()*9000)}`,
      now.toISOString()
    ]);
  }
  return rows;
}

function initDefaultTargets(sheet) {
  const rows = [];
  const year = new Date().getFullYear();
  const now = new Date();
  
  const add = (month, type, category, metric, val) => {
    rows.push([Utilities.getUuid(), year, month, type, category, metric, val, now]);
  };

  add(0, 'Main', 'Total', 'Sales', 120000000); 
  for(let m=1; m<=12; m++) {
    const v = (m === 3 || m === 9 || m === 12) ? 15000000 : 8000000;
    add(m, 'Main', 'Total', 'Sales', v);
  }

  const genres = [
    { name: '製薬：治験', channels: ['CMIC', 'その他CRO紹介', '既存顧客', '展示会', 'セミナー・ウェビナー'] },
    { name: '製薬：マーケティング', channels: ['CNZ連携', '広告代理店連携', '既存顧客', '展示会', 'セミナー・ウェビナー'] },
    { name: '運輸', channels: ['販売代理店連携', '既存顧客', '展示会', 'セミナー・ウェビナー'] }
  ];

  genres.forEach(g => {
    add(0, 'Genre', g.name, 'Sales', 40000000);
    g.channels.forEach(ch => {
      add(0, 'Channel', ch, 'Sales', 8000000);
      add(0, 'Channel', ch, 'Orders', 4);
      add(0, 'Channel', ch, 'Quotes', 10);
      add(0, 'Channel', ch, 'Meetings', 20);
      add(0, 'Channel', ch, 'FirstApo', 15);
      
      for(let m=1; m<=12; m++) {
        add(m, 'Channel', ch, 'Orders', 1);
        add(m, 'Channel', ch, 'Quotes', 2);
        add(m, 'Channel', ch, 'Meetings', 3);
        add(m, 'Channel', ch, 'FirstApo', 2);
      }
    });
  });

  if (rows.length > 0) sheet.getRange(2, 1, rows.length, 8).setValues(rows);
}

function generateDummyDeals(count) {
  const statuses = ['lead', 'topic', 'first_apo', 'needs', 'proposal', 'closing', 'pic_agree', 'contract', 'lost', 'vanished'];
  const engagementStatuses = ['contacting', 'dormant', 'out_of_scope'];
  const clients = ['株式会社アルファ', 'ベータ商事', 'ガンマ製薬', 'デルタ物流', 'オメガ・マーケティング', 'シグマ工業', 'ゼータ・ソリューションズ', 'イプシロン医療'];
  const genres = [
    { name: '製薬：治験', services: ['リクルーティングサービス', 'その他サービス'], channels: ['CMIC', '展示会'] },
    { name: '製薬：マーケティング', services: ['harmoPCS', 'SmartPDCA', 'ターゲティングサービス', 'その他マーケ向けサービス'], channels: ['広告代理店連携', 'CNZ連携'] },
    { name: '運輸', services: ['well-harmoレポート', 'harmo for Driver'], channels: ['既存顧客', '販売代理店連携'] }
  ];
  const rows = [];
  const now = new Date();
  for (let i = 0; i < count; i++) {
    const g = genres[Math.floor(Math.random() * genres.length)];
    const s = g.services[Math.floor(Math.random() * g.services.length)];
    const c = g.channels[Math.floor(Math.random() * g.channels.length)];
    const cl = clients[Math.floor(Math.random() * clients.length)];
    const monthOffset = Math.floor(Math.random() * 6) - 2;
    const d = new Date(now.getFullYear(), now.getMonth() + monthOffset, Math.floor(Math.random() * 28) + 1);
    const st = statuses[Math.floor(Math.random() * statuses.length)];
    const est = engagementStatuses[Math.floor(Math.random() * engagementStatuses.length)];
    const qDate = (['proposal', 'closing', 'pic_agree', 'contract'].includes(st)) ? formatDateStr(new Date(d.getTime()+7*86400000)) : '';
    const billingType = Math.random() > 0.75 ? 'monthly' : 'single';
    const billingDate = formatDateStr(new Date(d.getTime() + 3 * 86400000));
    const revenueMonth = formatDateStr(d).slice(0, 7);
    const billingStartMonth = billingType === 'single' ? '' : revenueMonth;
    const billingEndMonth = billingType === 'single' ? '' : revenueMonth;
    const billingTotal = (Math.floor(Math.random() * 50) + 10) * 100000;
    const billingAmount = billingType === 'single' ? billingTotal : Math.round(billingTotal / 3);

    rows.push([
      Utilities.getUuid(),
      cl,
      `${s}案件 ${i+1}`,
      (Math.floor(Math.random() * 50) + 10) * 100000,
      formatDateStr(d),
      qDate,
      st,
      est,
      g.name,
      s,
      c,
      billingType,
      billingDate,
      revenueMonth,
      billingStartMonth,
      billingEndMonth,
      billingAmount,
      billingTotal,
      new Date()
    ]);
  }
  return rows;
}

function generateDummyDealLogs(dealIds) {
  const logs = [];
  const now = new Date();
  dealIds.forEach(dealId => {
    const count = Math.floor(Math.random() * 4) + 1; 
    for (let i = 0; i < count; i++) {
      const r = Math.random();
      let type = 'NextAction';
      let content = '次回アクション予定';
      if (r < 0.3) { type = 'Meeting'; content = 'Web会議にて要件定義実施'; } 
      else if (r < 0.5) { type = 'Result'; content = '先方担当者より好感触を得た'; }

      const dateOffset = Math.floor(Math.random() * 60);
      const logDate = new Date(now.getTime() - dateOffset * 24 * 60 * 60 * 1000);
      logs.push([Utilities.getUuid(), dealId, formatDateStr(logDate), type, content, new Date()]);
    }
  });
  return logs;
}

function generateDummyActions(count) {
  const acts = ['ヒアリング', '提案書作成', '見積提出', 'クロージング'];
  const chs = ['CMIC', '既存顧客', '展示会'];
  const rows = [];
  const now = new Date();
  for (let i = 0; i < count; i++) {
    const dOff = Math.floor(Math.random() * 14) - 7;
    rows.push([
      Utilities.getUuid(),
      formatDateStr(new Date(now.getTime() + dOff * 86400000)),
      chs[Math.floor(Math.random() * chs.length)],
      acts[Math.floor(Math.random() * acts.length)],
      Math.random() > 0.7,
      '',
      '',
      'channel',
      new Date()
    ]);
  }
  return rows;
}

function generateDummyContents() {
  const now = new Date();
  return [
    [Utilities.getUuid(), '年末キャンペーンのご案内', 'キャンペーン', '年末限定の特別プランと申込特典を案内', '年末期間限定で初期費用を20%オフにするキャンペーンを実施しています。導入をご検討中の企業様向けに、無料相談枠も拡充しました。', '12/1〜12/31', '無料相談の希望日をお知らせください', now.toISOString()],
    [Utilities.getUuid(), '最新事例：治験リクルーティング効率化', '導入事例', '治験被験者募集のリードタイム短縮事例を紹介', '新規治験でリードタイムを30%短縮した成功事例をご紹介できます。既存の応募導線を整理し、対象者へのリーチを最適化しました。', '通年', '事例資料の送付をご希望ですか？', now.toISOString()],
    [Utilities.getUuid(), '2024年版 医療業界マーケティングトレンド', 'ホワイトペーパー', '医療業界のデジタル施策トレンドをまとめた資料', '医療業界のデジタル施策動向を網羅したホワイトペーパーをご案内可能です。社内共有用のPDFも用意しています。', '通年', '資料送付先のメールアドレスをご共有ください', now.toISOString()],
    [Utilities.getUuid(), '物流DX支援パッケージのご案内', 'サービス案内', '運輸向けDX支援パッケージの概要', '運輸業向けに現場KPI可視化と改善支援をセットにしたDXパッケージを提供しています。導入初期の伴走支援も含みます。', '通年', 'お打ち合わせ候補日を2〜3ついただけますか', now.toISOString()]
  ];
}

function initDefaultSettings(sheet) {
  const defaults = [
    ['Config', 'fiscalStartMonth', '10'],
    ['Config', 'fiscalYear', String(new Date().getFullYear())],
    ['Genre', '', '製薬：治験'], ['Genre', '', '製薬：マーケティング'], ['Genre', '', '運輸'],
    ['Service', '製薬：治験', 'リクルーティングサービス'], ['Service', '製薬：治験', 'その他サービス'],
    ['Channel', '製薬：治験', 'CMIC'], ['Channel', '製薬：治験', 'その他CRO紹介'], ['Channel', '製薬：治験', '既存顧客'], ['Channel', '製薬：治験', '展示会'], ['Channel', '製薬：治験', 'セミナー・ウェビナー'],
    ['Service', '製薬：マーケティング', 'harmoPCS'], ['Service', '製薬：マーケティング', 'SmartPDCA'], ['Service', '製薬：マーケティング', 'ターゲティングサービス'], ['Service', '製薬：マーケティング', 'その他マーケ向けサービス'],
    ['Channel', '製薬：マーケティング', 'CNZ連携'], ['Channel', '製薬：マーケティング', '広告代理店連携'], ['Channel', '製薬：マーケティング', '既存顧客'], ['Channel', '製薬：マーケティング', '展示会'], ['Channel', '製薬：マーケティング', 'セミナー・ウェビナー'],
    ['Service', '運輸', 'well-harmoレポート'], ['Service', '運輸', 'harmo for Driver'],
    ['Channel', '運輸', '販売代理店連携'], ['Channel', '運輸', '既存顧客'], ['Channel', '運輸', '展示会'], ['Channel', '運輸', 'セミナー・ウェビナー'],
  ];
  const rows = defaults.map(d => [Utilities.getUuid(), d[0], d[1], d[2], new Date()]);
  sheet.getRange(2, 1, rows.length, 5).setValues(rows);
}

// --- CRUD Functions ---
function saveDeal(deal) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DEALS);
  if (sheet) ensureDealsSchema(sheet);

  return saveRow(SHEET_DEALS, deal, ['id', 'clientName', 'dealName', 'amount', 'occurrenceDate', 'quoteDate', 'status', 'engagementStatus', 'genre', 'service', 'channel', 'billingType', 'billingDate', 'revenueMonth', 'billingStartMonth', 'billingEndMonth', 'billingAmount', 'billingTotal', 'contactId', 'contactIds', 'updatedAt'], ['clientName']);
}
function deleteDeal(id) {
  const deleted = deleteRowById(SHEET_DEALS, id);
  if (deleted) deleteRowsByColumnValue(SHEET_DEAL_LOGS, 'dealId', id);
  return deleted;
}
function saveAction(action) {
  if (!action.actionContent) throw new Error("内容を入力してください");
  action.isDone = action.isDone ? true : false;

  // Schema Fix: Ensure 'genre' header exists
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ACTIONS);
  ensureActionsSchema(sheet);
  const lastCol = sheet.getLastColumn();
  if (lastCol > 0) { // Check if sheet has headers
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (!headers.includes('genre')) {
      sheet.getRange(1, lastCol + 1).setValue('genre').setFontWeight('bold');
    }
  }

  // Use appended column order to preserve existing data structure
  return saveRow(SHEET_ACTIONS, action, ['id', 'targetDate', 'channel', 'actionContent', 'isDone', 'genre', 'dealId', 'actionType', 'updatedAt'], ['actionContent']);
}
function deleteAction(id) { return deleteRowById(SHEET_ACTIONS, id); }
function toggleActionStatus(id, isDone) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ACTIONS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false; // No data
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const idx = ids.indexOf(id);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const isDoneIdx = headers.indexOf('isDone');
  if (idx !== -1 && isDoneIdx !== -1) { sheet.getRange(idx + 2, isDoneIdx + 1).setValue(isDone); return true; }
  return false;
}
function saveSetting(item) { return saveRow(SHEET_SETTINGS, item, ['id', 'type', 'category', 'name', 'updatedAt']); }
function deleteSetting(id) { return deleteRowById(SHEET_SETTINGS, id); }
function saveTarget(target) { return saveRow(SHEET_TARGETS, target, ['id', 'year', 'month', 'type', 'category', 'metric', 'value', 'updatedAt']); }
function saveDealLog(log) { return saveRow(SHEET_DEAL_LOGS, log, ['id', 'dealId', 'logDate', 'logType', 'content', 'updatedAt'], ['dealId', 'content']); }
function saveDealLog(log) { return saveRow(SHEET_DEAL_LOGS, log, ['id', 'dealId', 'logDate', 'logType', 'content', 'updatedAt'], ['dealId', 'content']); }
function deleteDealLog(id) { return deleteRowById(SHEET_DEAL_LOGS, id); }
function saveContact(contact) { return saveRow(SHEET_CONTACTS, contact, ['id', 'company', 'name', 'department', 'email', 'mobilePhone', 'updatedAt'], ['name']); }
function deleteContact(id) { return deleteRowById(SHEET_CONTACTS, id); }
function saveContent(content) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet(ss, SHEET_CONTENTS, ['id', 'title', 'category', 'summary', 'body', 'campaignPeriod', 'cta', 'updatedAt']);
  ensureContentsSchema(ss.getSheetByName(SHEET_CONTENTS));
  return saveRow(SHEET_CONTENTS, content, ['id', 'title', 'category', 'summary', 'body', 'campaignPeriod', 'cta', 'updatedAt'], ['title']);
}
function deleteContent(id) { return deleteRowById(SHEET_CONTENTS, id); }
function saveBillingMaster(item) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet(ss, SHEET_BILLING_MASTER, BILLING_MASTER_HEADERS);
  ensureBillingMasterSchema(ss.getSheetByName(SHEET_BILLING_MASTER));
  if (!item.id && item.dealId) item.id = item.dealId;
  item.contractAmount = Number(item.contractAmount || 0);
  if (item.transactionType && item.transactionType !== '新規') {
    item.creditCheck = '－';
  }
  return saveRow(SHEET_BILLING_MASTER, item, BILLING_MASTER_HEADERS, ['id']);
}
function saveMonthlyTask(task) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet(ss, SHEET_MONTHLY_TASKS, MONTHLY_TASK_HEADERS);
  ensureMonthlyTasksSchema(ss.getSheetByName(SHEET_MONTHLY_TASKS));
  return saveRow(SHEET_MONTHLY_TASKS, task, MONTHLY_TASK_HEADERS, ['targetMonth', 'dealId', 'taskType']);
}
function generateMonthlyTasks(targetMonth) {
  if (!targetMonth) throw new Error('対象年月が未指定です');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet(ss, SHEET_BILLING_MASTER, BILLING_MASTER_HEADERS);
  ensureSheet(ss, SHEET_MONTHLY_TASKS, MONTHLY_TASK_HEADERS);
  ensureBillingMasterSchema(ss.getSheetByName(SHEET_BILLING_MASTER));
  ensureMonthlyTasksSchema(ss.getSheetByName(SHEET_MONTHLY_TASKS));
  ensureBillingMasterRowsFromDeals(ss);

  const masters = getSheetData(ss, SHEET_BILLING_MASTER);
  const existingTasks = getSheetData(ss, SHEET_MONTHLY_TASKS);
  const taskMap = new Map();
  existingTasks.forEach(task => {
    const key = `${task.targetMonth}__${task.dealId}__${task.taskType}`;
    taskMap.set(key, task);
  });

  const isSemiAnnualMonth = (monthStr) => {
    const ym = normalizeYearMonth(monthStr);
    if (!ym) return false;
    const m = Number(ym.split('-')[1]);
    return m === 3 || m === 9;
  };

  const upsertTask = (dealId, taskType, notes, boxUrl) => {
    const key = `${targetMonth}__${dealId}__${taskType}`;
    const existing = taskMap.get(key);
    const payload = existing ? { ...existing } : {
      id: Utilities.getUuid(),
      targetMonth,
      dealId,
      taskType,
      status: '未',
      saifuStatus: '未',
      notes: '',
      boxUrl: ''
    };
    if (notes !== undefined) payload.notes = notes || '';
    if (boxUrl !== undefined) payload.boxUrl = boxUrl || '';
    saveRow(SHEET_MONTHLY_TASKS, payload, MONTHLY_TASK_HEADERS, ['targetMonth', 'dealId', 'taskType']);
  };

  const prevMonth = getPreviousMonth(targetMonth);
  masters.forEach(master => {
    const dealId = master.id;
    if (!dealId) return;
    const start = normalizeYearMonth(master.contractStartMonth);
    const end = normalizeYearMonth(master.contractEndMonth || master.contractStartMonth);
    if (!start) return;

    if (isMonthInRange(targetMonth, start, end)) {
      upsertTask(dealId, '売上請求', master.billingType || '', master.boxUrl || '');
    }

    const genre = master.genre || '';
    const allianceDataUsage = master.allianceDataUsage || '無';
    if (genre.includes('製薬') && allianceDataUsage === '有') {
      if (isMonthInRange(prevMonth, start, end)) {
        upsertTask(dealId, '支払明細', '', master.boxUrl || '');
      }
    }

    if (genre === '運輸' && master.salesFlow === 'パートナー') {
      if (isSemiAnnualMonth(targetMonth) && isMonthInRange(targetMonth, start, end)) {
        upsertTask(dealId, '手数料支払', master.paymentPattern || '', master.boxUrl || '');
      }
    }
  });

  return getSheetData(ss, SHEET_MONTHLY_TASKS).filter(task => task.targetMonth === targetMonth);
}

// --- Helpers ---
function saveRow(sheetName, data, columns, requiredFields = []) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  for (const field of requiredFields) if (!data[field]) throw new Error(`Missing: ${field}`);
  if (!data.id) data.id = Utilities.getUuid();
  
  // Date型をISO文字列に変換して保存する前に、オブジェクト内のDateを文字列化する
  const now = new Date();
  data.updatedAt = now.toISOString(); // 文字列として保存

  const lastRow = sheet.getLastRow();
  const ids = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat() : [];
  const index = ids.indexOf(data.id);
  const rowToUpdate = index !== -1 ? index + 2 : lastRow + 1;
  const rowData = columns.map(col => col === 'amount' ? Number(data[col]) : data[col]);
  sheet.getRange(rowToUpdate, 1, 1, rowData.length).setValues([rowData]);
  return data;
}
function deleteRowById(sheetName, id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const index = ids.indexOf(id);
  if (index !== -1) { sheet.deleteRow(index + 2); return true; }
  return false;
}

function deleteRowsByColumnValue(sheetName, columnName, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const columnIndex = headers.indexOf(columnName);
  if (columnIndex === -1) return 0;

  const values = sheet.getRange(2, columnIndex + 1, lastRow - 1, 1).getValues().flat();
  let deletedCount = 0;

  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i] === value) {
      sheet.deleteRow(i + 2);
      deletedCount++;
    }
  }

  return deletedCount;
}
function resetSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) sheet.clear(); else sheet = ss.insertSheet(sheetName);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  return sheet;
}
function findSheetByName(ss, sheetName) {
  const direct = ss.getSheetByName(sheetName);
  if (direct) return direct;
  const normalized = String(sheetName || '').trim().toLowerCase();
  const match = ss.getSheets().find(s => String(s.getName() || '').trim().toLowerCase() === normalized);
  return match || null;
}
function ensureSheet(ss, sheetName, headers) {
  let sheet = findSheetByName(ss, sheetName);
  if (!sheet) {
    try {
      sheet = ss.insertSheet(sheetName);
    } catch (error) {
      sheet = findSheetByName(ss, sheetName);
    }
  }
  if (sheet && sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }
  return sheet;
}

// Migration helper to apply latest Deals schema without altering existing data
function backfillEngagementStatusColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DEALS);
  if (!sheet) return false;
  ensureDealsSchema(sheet);
  return true;
}

// Ensure Deals sheet has required columns (billing + contact linkage included)
function ensureDealsSchema(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;

  const refreshHeaders = () => sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const expected = [
    'id',
    'clientName',
    'dealName',
    'amount',
    'occurrenceDate',
    'quoteDate',
    'status',
    'engagementStatus',
    'genre',
    'service',
    'channel',
    'billingType',
    'billingDate',
    'revenueMonth',
    'billingStartMonth',
    'billingEndMonth',
    'billingAmount',
    'billingTotal',
    'contactId',
    'contactIds',
    'updatedAt'
  ];

  let headers = refreshHeaders();
  const lastRow = sheet.getLastRow();
  let insertedContactIds = false;

  const insertColumn = (name) => {
    if (headers.includes(name)) return;
    const nameIdx = expected.indexOf(name);
    let insertCol = sheet.getLastColumn() + 1;
    for (let i = nameIdx + 1; i < expected.length; i++) {
      const nextIdx = headers.indexOf(expected[i]);
      if (nextIdx !== -1) {
        insertCol = nextIdx + 1;
        break;
      }
    }
    sheet.insertColumnBefore(insertCol);
    sheet.getRange(1, insertCol).setValue(name).setFontWeight('bold');
    headers = refreshHeaders();
  };

  expected.forEach(col => {
    const wasMissing = !headers.includes(col);
    insertColumn(col);
    if (col === 'engagementStatus' && wasMissing && lastRow > 1) {
      const engagementStatusIdx = headers.indexOf('engagementStatus');
      if (engagementStatusIdx !== -1) {
        const defaults = Array.from({ length: lastRow - 1 }, () => ['contacting']);
        sheet.getRange(2, engagementStatusIdx + 1, lastRow - 1, 1).setValues(defaults);
      }
    }
    if (col === 'contactIds' && wasMissing) {
      insertedContactIds = true;
    }
  });

  if (insertedContactIds && lastRow > 1) {
    const contactIdIdx = headers.indexOf('contactId');
    const contactIdsIdx = headers.indexOf('contactIds');
    if (contactIdIdx !== -1 && contactIdsIdx !== -1) {
      const values = sheet.getRange(2, contactIdIdx + 1, lastRow - 1, 1).getValues().flat();
      const migrated = values.map(v => [v || '']);
      sheet.getRange(2, contactIdsIdx + 1, lastRow - 1, 1).setValues(migrated);
    }
  }
}
function ensureActionsSchema(sheet) {
  if (!sheet) return;
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;

  const refreshHeaders = () => sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let headers = refreshHeaders();
  let updatedAtIdx = headers.indexOf('updatedAt');

  // Ensure isDone column
  if (!headers.includes('isDone')) {
    let targetCol;
    if (updatedAtIdx !== -1) {
      targetCol = updatedAtIdx + 1;
      sheet.insertColumnBefore(targetCol);
    } else {
      sheet.insertColumnAfter(sheet.getLastColumn());
      targetCol = sheet.getLastColumn();
    }
    sheet.getRange(1, targetCol).setValue('isDone').setFontWeight('bold');
    headers = refreshHeaders();
    updatedAtIdx = headers.indexOf('updatedAt');
    const isDoneIdx = headers.indexOf('isDone');
    const lastRow = sheet.getLastRow();
    if (lastRow > 1 && isDoneIdx !== -1) {
      const defaults = Array.from({ length: lastRow - 1 }, () => [false]);
      sheet.getRange(2, isDoneIdx + 1, lastRow - 1, 1).setValues(defaults);
    }
  }

  // Ensure genre column
  if (!headers.includes('genre')) {
    let targetCol;
    if (updatedAtIdx !== -1) {
      targetCol = updatedAtIdx + 1; // 1-based position of updatedAt
      sheet.insertColumnBefore(targetCol);
    } else {
      sheet.insertColumnAfter(sheet.getLastColumn());
      targetCol = sheet.getLastColumn();
    }
    sheet.getRange(1, targetCol).setValue('genre').setFontWeight('bold');
    headers = refreshHeaders();
    updatedAtIdx = headers.indexOf('updatedAt');
  }

  // Ensure dealId column (for per-deal action timeline)
  if (!headers.includes('dealId')) {
    let targetCol;
    if (updatedAtIdx !== -1) {
      targetCol = updatedAtIdx + 1; // before updatedAt
      sheet.insertColumnBefore(targetCol);
    } else {
      sheet.insertColumnAfter(sheet.getLastColumn());
      targetCol = sheet.getLastColumn();
    }
    sheet.getRange(1, targetCol).setValue('dealId').setFontWeight('bold');
  }

  // Ensure actionType column (deal vs channel)
  headers = refreshHeaders();
  updatedAtIdx = headers.indexOf('updatedAt');
  if (!headers.includes('actionType')) {
    let targetCol;
    if (updatedAtIdx !== -1) {
      targetCol = updatedAtIdx + 1;
      sheet.insertColumnBefore(targetCol);
    } else {
      sheet.insertColumnAfter(sheet.getLastColumn());
      targetCol = sheet.getLastColumn();
    }
    sheet.getRange(1, targetCol).setValue('actionType').setFontWeight('bold');
    headers = refreshHeaders();
    const actionTypeIdx = headers.indexOf('actionType');
    const dealIdIdx = headers.indexOf('dealId');
    const lastRow = sheet.getLastRow();
    if (lastRow > 1 && actionTypeIdx !== -1) {
      const dealIds = dealIdIdx !== -1 ? sheet.getRange(2, dealIdIdx + 1, lastRow - 1, 1).getValues().flat() : [];
      const defaults = dealIds.map(id => [id ? 'deal' : 'channel']);
      sheet.getRange(2, actionTypeIdx + 1, lastRow - 1, 1).setValues(defaults);
    }
  }
}

function ensureContactsSchema(sheet) {
  if (!sheet) return;
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const mobileIdx = headers.indexOf('mobilePhone');
  const phoneIdx = headers.indexOf('phone');
  const updatedAtIdx = headers.indexOf('updatedAt');

  if (mobileIdx === -1) {
    if (phoneIdx !== -1) {
      sheet.getRange(1, phoneIdx + 1).setValue('mobilePhone').setFontWeight('bold');
    } else {
      const insertCol = updatedAtIdx !== -1 ? updatedAtIdx + 1 : sheet.getLastColumn() + 1;
      sheet.insertColumnBefore(insertCol);
      sheet.getRange(1, insertCol).setValue('mobilePhone').setFontWeight('bold');
    }
  }
}

function ensureContentsSchema(sheet) {
  if (!sheet) return;
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;

  const expected = ['id', 'title', 'category', 'summary', 'body', 'campaignPeriod', 'cta', 'updatedAt'];
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  expected.forEach((col) => {
    if (!headers.includes(col)) {
      sheet.insertColumnAfter(sheet.getLastColumn());
      sheet.getRange(1, sheet.getLastColumn()).setValue(col).setFontWeight('bold');
    }
  });
}

function ensureSheetSchema(sheet, expected) {
  if (!sheet) return;
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;

  const refreshHeaders = () => sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let headers = refreshHeaders();

  const insertColumn = (name) => {
    if (headers.includes(name)) return;
    const nameIdx = expected.indexOf(name);
    let insertCol = sheet.getLastColumn() + 1;
    for (let i = nameIdx + 1; i < expected.length; i++) {
      const nextIdx = headers.indexOf(expected[i]);
      if (nextIdx !== -1) {
        insertCol = nextIdx + 1;
        break;
      }
    }
    sheet.insertColumnBefore(insertCol);
    sheet.getRange(1, insertCol).setValue(name).setFontWeight('bold');
    headers = refreshHeaders();
  };

  expected.forEach((col) => insertColumn(col));
}

function ensureBillingMasterSchema(sheet) {
  ensureSheetSchema(sheet, BILLING_MASTER_HEADERS);
}

function ensureMonthlyTasksSchema(sheet) {
  ensureSheetSchema(sheet, MONTHLY_TASK_HEADERS);
}

// Rebuild header rows without altering sheet data.
function rebuildSheetHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headerMap = {
    [SHEET_SETTINGS]: ['id', 'type', 'category', 'name', 'updatedAt'],
    [SHEET_DEALS]: ['id', 'clientName', 'dealName', 'amount', 'occurrenceDate', 'quoteDate', 'status', 'engagementStatus', 'genre', 'service', 'channel', 'billingType', 'billingDate', 'revenueMonth', 'billingStartMonth', 'billingEndMonth', 'billingAmount', 'billingTotal', 'contactId', 'contactIds', 'updatedAt'],
    [SHEET_ACTIONS]: ['id', 'targetDate', 'channel', 'actionContent', 'isDone', 'genre', 'dealId', 'actionType', 'updatedAt'],
    [SHEET_TARGETS]: ['id', 'year', 'month', 'type', 'category', 'metric', 'value', 'updatedAt'],
    [SHEET_DEAL_LOGS]: ['id', 'dealId', 'logDate', 'logType', 'content', 'updatedAt'],
    [SHEET_CONTACTS]: ['id', 'company', 'name', 'department', 'email', 'mobilePhone', 'updatedAt'],
    [SHEET_CONTENTS]: ['id', 'title', 'category', 'summary', 'body', 'campaignPeriod', 'cta', 'updatedAt'],
    [SHEET_BILLING_MASTER]: BILLING_MASTER_HEADERS,
    [SHEET_MONTHLY_TASKS]: MONTHLY_TASK_HEADERS
  };

  Object.entries(headerMap).forEach(([sheetName, expected]) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return;

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const updated = [...headers];

    headers.forEach((header, idx) => {
      if (sheetName === SHEET_CONTACTS && header === 'phone') {
        updated[idx] = 'mobilePhone';
        return;
      }
      if (expected.includes(header)) updated[idx] = header;
    });

    if (updated.some((v, i) => v !== headers[i])) {
      sheet.getRange(1, 1, 1, lastCol).setValues([updated]).setFontWeight('bold');
    }
  });
}
function getSheetData(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  return data.map(row => {
    let obj = {};
    headers.forEach((header, i) => {
      if (row[i] instanceof Date) obj[header] = formatDateStr(row[i]);
      else obj[header] = row[i];
    });
    return obj;
  });
}
function formatDateStr(date) { return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function normalizeYearMonth(value) {
  if (!value) return '';
  if (value instanceof Date) return formatDateStr(value).slice(0, 7);
  const str = String(value).trim();
  if (/^\d{4}-\d{2}$/.test(str)) return str;
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) return str.slice(0, 7);
  if (/^\d{4}\/\d{2}/.test(str)) return str.slice(0, 7).replace(/\//g, '-');
  return '';
}
function getPreviousMonth(targetMonth) {
  const ym = normalizeYearMonth(targetMonth);
  if (!ym) return '';
  const [y, m] = ym.split('-').map(Number);
  const date = new Date(y, m - 2, 1);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
}
function isMonthInRange(targetMonth, startMonth, endMonth) {
  const target = normalizeYearMonth(targetMonth);
  const start = normalizeYearMonth(startMonth);
  const end = normalizeYearMonth(endMonth || startMonth);
  if (!target || !start) return false;
  const rangeEnd = end || start;
  return target >= start && target <= rangeEnd;
}

function ensureBillingMasterRowsFromDeals(ss) {
  const masterSheet = ss.getSheetByName(SHEET_BILLING_MASTER);
  const dealsSheet = ss.getSheetByName(SHEET_DEALS);
  if (!masterSheet || !dealsSheet) return;
  const masters = getSheetData(ss, SHEET_BILLING_MASTER);
  const masterIds = new Set(masters.map(m => m.id));
  const deals = getSheetData(ss, SHEET_DEALS);

  deals.forEach(deal => {
    if (!deal?.id || masterIds.has(deal.id)) return;
    const startMonth = normalizeYearMonth(deal.billingStartMonth) || normalizeYearMonth(deal.revenueMonth) || normalizeYearMonth(deal.occurrenceDate);
    const endMonth = normalizeYearMonth(deal.billingEndMonth) || startMonth;
    const payload = {
      id: deal.id,
      dealName: deal.dealName || '',
      clientName: deal.clientName || '',
      genre: deal.genre || '',
      contractAmount: Number(deal.amount || 0),
      contractStartMonth: startMonth || '',
      contractEndMonth: endMonth || '',
      billingType: deal.billingType || '',
      pjNumber: '',
      transactionType: '',
      creditCheck: '未',
      boxUrl: '',
      saifuStatus: '未',
      deliveryStatus: '未',
      invoiceStatus: '未',
      salesFlow: '',
      allianceDataUsage: '無',
      paymentPartner: '',
      paymentPattern: ''
    };
    saveRow(SHEET_BILLING_MASTER, payload, BILLING_MASTER_HEADERS, []);
  });
}

// --- AI (Gemini) Functions ---

function getGeminiApiKey() {
  // Use Script Properties for security
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
}

function callGemini(prompt, systemInstruction = null) {
  const apiKey = getGeminiApiKey();
  if (!apiKey) throw new Error("API Keyが設定されていません。GASエディタの「プロジェクトの設定 > スクリプトプロパティ」に 'GEMINI_API_KEY' を追加してください。");

  const model = 'gemini-3-flash-preview'; // User explicitly requested this model
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  const contents = [{ role: "user", parts: [{ text: prompt }] }];
  if (systemInstruction) {
    // Note: v1beta support for system_instruction might vary by endpoint, 
    // but putting it in the payload is standard for newer Gemini models.
    // If simple prompt is preferred, we can prepend it.
    // For now, let's prepend to prompt to be safe across versions unless using specific API structure.
    contents[0].parts[0].text = `System Instruction: ${systemInstruction}\n\nUser Input: ${prompt}`;
  }

  const payload = {
    contents: contents,
    generationConfig: {
      temperature: 0.2, // Low temperature for extraction tasks
      responseMimeType: "application/json" // Force JSON response
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const shouldRetry = (json, statusCode) => {
    if (statusCode === 429 || statusCode === 503) return true;
    const message = json?.error?.message || '';
    return message.includes('overloaded') || message.includes('try again later');
  };

  const maxAttempts = 3;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    const json = JSON.parse(response.getContentText());

    if (json.error) {
      if (attempt < maxAttempts && shouldRetry(json, statusCode)) {
        Utilities.sleep(1000 * attempt);
        continue;
      }
      throw new Error(`Gemini API Error: ${json.error.message}`);
    }

    try {
      let contentText = json.candidates[0].content.parts[0].text;
      // Cleanup markdown code blocks if present
      contentText = contentText.replace(/^```json\s*/, "").replace(/^```\s*/, "").replace(/\s*```$/, "");
      return JSON.parse(contentText);
    } catch (e) {
      throw new Error("AIからの応答の解析に失敗しました: " + e.message + " Response: " + (json.candidates?.[0]?.content?.parts?.[0]?.text || "No content"));
    }
  }

  throw new Error('Gemini API Error: モデルが混雑しています。時間をおいて再度お試しください。');
}

function processContactText(text) {
  const systemPrompt = `
    あなたは連絡先情報の抽出AIです。
    入力テキストから、以下の情報を抽出し、JSON形式で返してください。
    - company: 会社名 (株式会社などはそのまま)
    - name: 氏名 (フルネーム), 不明な場合は空文字
    - department: 部署・役職, 不明な場合は空文字
    - email: メールアドレス, 不明な場合は空文字
    - mobilePhone: 携帯電話番号, 不明な場合は空文字
    
    JSONのキーは必ず上記を使用してください。
  `;
  return callGemini(text, systemPrompt);
}

function processMeetingText(dealId, text) {
  const now = new Date();
  const todayStr = formatDateStr(now);

  const systemPrompt = `
    あなたは優秀なビジネスアシスタントです。
    会議のメモや議事録から、重要なポイントを抽出し、以下の構造のJSONで返してください。

    # 前提条件
    - 現在のシステム日付は「${todayStr}」です。
    - 入力テキストに「年」が明記されていない日付（例：「12月1日」）は、必ずシステム日付の年（${now.getFullYear()}年）またはその翌年として解釈してください。
    - 【重要】文脈がない限り、過去の年（2023年や2024年など）には絶対にしないでください。
    - 「明日」や「来週」などの相対日時は、基準日「${todayStr}」から計算してください。

    # 出力形式 (JSON)
    {
      "summary": "会議の要約（50文字以内で簡潔に。常体で記述）",
      "date": "会議実施日（YYYY-MM-DD形式）。文脈から判断できない場合は基準日(${todayStr})",
      "next_actions": [
        "具体的なネクストアクション1（主語と期限を含める）",
        "具体的なネクストアクション2"
      ]
    }
    
    ネクストアクションがない場合は空配列にしてください。
  `;

  const result = callGemini(text, systemPrompt);
  
  // Save to DB directly
  const savedLogs = [];
  // const now = new Date(); // Moved up
  
  // 1. Save Meeting Log
  const mtgLog = {
    id: Utilities.getUuid(),
    dealId: dealId,
    logDate: result.date || formatDateStr(now),
    logType: 'Meeting',
    content: result.summary,
    updatedAt: now.toISOString()
  };
  saveDealLog(mtgLog);
  savedLogs.push(mtgLog);

  // 2. Save Next Actions
  if (result.next_actions && Array.isArray(result.next_actions)) {
    result.next_actions.forEach(actionContent => {
      const actLog = {
        id: Utilities.getUuid(),
        dealId: dealId, // Link to Deal
        logDate: result.date || formatDateStr(now), // Use same date or calculate deadline if possible
        logType: 'NextAction',
        content: actionContent,
        updatedAt: now.toISOString()
      };
      saveDealLog(actLog);
      savedLogs.push(actLog);
    });
  }

  return { logs: savedLogs, systemDate: todayStr, aiDate: result.date }; // Return object wrapper
}

function generateReengagementEmail(payload) {
  const dealId = payload?.dealId;
  if (!dealId) throw new Error('dealIdが必要です');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const deals = getSheetData(ss, SHEET_DEALS);
  const logs = getSheetData(ss, SHEET_DEAL_LOGS);
  const contacts = getSheetData(ss, SHEET_CONTACTS);
  const contents = getSheetData(ss, SHEET_CONTENTS);

  const deal = deals.find(d => d.id === dealId);
  if (!deal) throw new Error('案件が見つかりません');

  const contact = contacts.find(c => c.id === payload?.contactId) || null;
  const selectedContents = (payload?.contentIds || []).map(id => contents.find(c => c.id === id)).filter(Boolean);

  const recentLogs = logs
    .filter(l => l.dealId === dealId && ['NextAction', 'Result', 'Meeting'].includes(l.logType))
    .sort((a, b) => String(b.logDate || '').localeCompare(String(a.logDate || '')))
    .slice(0, 6);

  const systemPrompt = `
    あなたはB2B営業の掘り起こしメール作成AIです。
    以下の情報を基に、丁寧で前向きな掘り起こしメールを日本語で作成してください。
    - 過去のやり取り (NextAction/Result/Meeting) を必ず踏まえる
    - 新しく伝えたいコンテンツ (content) の要点を自然に織り交ぜる
    - 相手の時間を尊重し、短めで読みやすい構成にする
    - 押し売りはしない、次の一歩を提案する
    出力は次のJSON形式のみ:
    {
      "subject": "件名",
      "body": "本文"
    }
  `;

  const prompt = JSON.stringify({
    deal,
    contact,
    recentLogs,
    content: selectedContents,
    note: payload?.note || ''
  }, null, 2);

  return callGemini(prompt, systemPrompt);
}

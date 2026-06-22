// ============================================================
// LINE家計簿システム - Google Apps Script
// Code.gs
// ============================================================

const LINE_CHANNEL_ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const MY_USER_ID = PropertiesService.getScriptProperties().getProperty('MY_USER_ID');
const GITHUB_TOKEN = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');

const SHEET_LOG = '記録';
const SHEET_CONFIG = '設定';
const SHEET_COMMANDS = 'コマンド';
const SHEET_INCOME = '収入';
const FIXED_CATEGORIES = ['固定費'];

function doPost(e) {
  if (!e || !e.postData) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.source === 'liff') {
      recordExpense(body.category, body.amount, body.memo, body.date);
      const report = getBalanceReport();
      pushMessage(body.userId,
        `✅ 記録しました！\n📁 ${body.category}：¥${Number(body.amount).toLocaleString()}\n\n${report}`
      );
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const events = body.events;
    events.forEach(event => {
      if (event.type === 'message' && event.message.type === 'text') {
        handleMessage(event.source.userId, event.message.text.trim(), event.replyToken);
      } else if (event.type === 'postback') {
        handlePostback(event.source.userId, event.postback.data, event.replyToken);
      }
    });
  } catch (err) {
    console.log('doPost エラー: ' + err.toString());
  }
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function callLineApi(url, payload) {
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    console.log('LINE API レスポンス: ' + response.getContentText());
  } catch (err) {
    console.log('LINE API エラー: ' + err.toString());
  }
}

function getMainQuickReplies() {
  return [
    { type: 'message', label: '💰 残高確認', text: '残高確認' },
    { type: 'message', label: '📊 今月の集計', text: '今月の集計' },
    { type: 'message', label: '📋 今月の明細', text: '今月の明細' },
    { type: 'message', label: '✏️ 記録する', text: '記録する' }
  ];
}

function handleMessage(userId, message, replyToken) {
  if (message === 'myid') {
    replyMessage(replyToken, userId);
    return;
  }

  const props = PropertiesService.getScriptProperties();
  const waitingCategory = props.getProperty('waitingCategory_' + userId);
  const waitingIncomeCategory = props.getProperty('waitingIncomeCategory_' + userId);
  const categories = getCategoryNames();
  const incomeCategories = getIncomeCategoryNames();

  const parts = message.split(/\s+/);

  // 収入の一行入力: 「給料 250000」
  if (parts.length >= 2 && incomeCategories.includes(parts[0]) && /^\d+$/.test(parts[parts.length - 1])) {
    const category = parts[0];
    const amount = parseInt(parts[parts.length - 1]);
    const memo = parts.slice(1, parts.length - 1).join(' ');
    recordIncome(category, amount, memo);
    props.deleteProperty('waitingIncomeCategory_' + userId);
    const label = memo ? `${category}（${memo}）` : category;
    replyMessage(replyToken,
      `✅ 収入を記録しました！\n💰 ${label}：${amount.toLocaleString()}円`,
      getMainQuickReplies()
    );
    return;
  }

  // 支出の一行入力: 「食費 コンビニ 400」
  if (parts.length >= 2 && categories.includes(parts[0]) && /^\d+$/.test(parts[parts.length - 1])) {
    const category = parts[0];
    const amount = parseInt(parts[parts.length - 1]);
    const memo = parts.slice(1, parts.length - 1).join(' ');
    recordExpense(category, amount, memo);
    props.deleteProperty('waitingCategory_' + userId);
    const report = getBalanceReport();
    const label = memo ? `${category}（${memo}）` : category;
    replyMessage(replyToken,
      `✅ 記録しました！\n📁 ${label}：${amount.toLocaleString()}円\n\n${report}`,
      getMainQuickReplies()
    );
    return;
  }

  // 収入カテゴリの金額入力待ち状態
  if (waitingIncomeCategory && /^\d+$/.test(message)) {
    const amount = parseInt(message, 10);
    recordIncome(waitingIncomeCategory, amount);
    props.deleteProperty('waitingIncomeCategory_' + userId);
    replyMessage(replyToken,
      `✅ 収入を記録しました！\n💰 ${waitingIncomeCategory}：${amount.toLocaleString()}円`,
      getMainQuickReplies()
    );
    return;
  }

  // 支出カテゴリの金額入力待ち状態
  if (waitingCategory && /^\d+$/.test(message)) {
    const amount = parseInt(message, 10);
    recordExpense(waitingCategory, amount);
    props.deleteProperty('waitingCategory_' + userId);
    const report = getBalanceReport();
    replyMessage(replyToken,
      `✅ 記録しました！\n📁 ${waitingCategory}：${amount.toLocaleString()}円\n\n${report}`,
      getMainQuickReplies()
    );
    return;
  }

  // 収入カテゴリ名のみ送信
  if (incomeCategories.includes(message)) {
    props.setProperty('waitingIncomeCategory_' + userId, message);
    replyMessage(replyToken, `💰 ${message}\n金額を数字で入力してください（例：250000）`);
    return;
  }

  // 支出カテゴリ名のみ送信
  if (categories.includes(message)) {
    props.setProperty('waitingCategory_' + userId, message);
    replyMessage(replyToken, `📁 ${message}\n金額を数字で入力してください（例：500）`);
    return;
  }

  if (message === '記録する') {
    replyMessage(replyToken, '📁 カテゴリを選んでください',
      categories.map(cat => ({ type: 'message', label: cat, text: cat }))
    );
    return;
  }

  if (message === '残高確認') { replyMessage(replyToken, getBalanceReport(), getMainQuickReplies()); return; }
  if (message === '今月の集計') { replyMessage(replyToken, getMonthlyReport(), getMainQuickReplies()); return; }
  if (message === '今月の明細') { replyMessage(replyToken, getMonthlyDetail(), getMainQuickReplies()); return; }
  if (message === '週末プラン') { replyMessage(replyToken, getWeekendPlanText()); return; }

  // ── QuickLog Bot ─────────────────────────────
  const weightMatch = message.match(/^体重[\s\n]+([\d.]+)/);
  if (weightMatch) {
    logWeight(weightMatch[1]);
    replyMessage(replyToken, '体重を記録しました ✅\n' + weightMatch[1] + 'kg（' + getDateStr_() + '）');
    return;
  }
  const memoMatch = message.match(/^メモ[\s\n]+(.+)/s);
  if (memoMatch) {
    saveMemo(memoMatch[1].trim());
    replyMessage(replyToken, 'メモしました ✅\n「' + memoMatch[1].trim() + '」');
    return;
  }
  if (message === '買い物リスト') {
    replyMessage(replyToken, getShoppingList());
    return;
  }
  const shoppingAddMatch = message.match(/^買い物\n(.+)/s);
  if (shoppingAddMatch) {
    addShoppingItem(shoppingAddMatch[1].trim());
    replyMessage(replyToken, '買い物リストに追加しました ✅\n・' + shoppingAddMatch[1].trim());
    return;
  }
  if (message === '体重一覧') {
    replyMessage(replyToken, getWeightLog());
    return;
  }

  if (message === 'ヘルプ' || message === 'help') {
    replyMessage(replyToken, getHelpMessage(), getMainQuickReplies());
    return;
  }

  const commandList = getCommandList();
  const matched = commandList.find(row => row[0] === message);
  if (matched && matched[2]) {
    replyMessage(replyToken, matched[2], getMainQuickReplies());
    return;
  }

  replyMessage(replyToken,
    '📋 使い方\n' +
    '① 一行入力：「食費 コンビニ 400」\n' +
    '② ボタン → 金額入力\n' +
    '③ 下のボタンから確認\n\n' +
    'カテゴリ：' + categories.join('、'),
    getMainQuickReplies()
  );
}

function handlePostback(userId, data, replyToken) {
  const params = parseQuery(data);
  if (params.action === 'category') {
    handleMessage(userId, params.value, replyToken);
  } else if (params.action === 'balance') {
    handleMessage(userId, '残高確認', replyToken);
  } else if (params.action === 'monthly') {
    handleMessage(userId, '今月の集計', replyToken);
  } else if (params.action === 'weekend_plan') {
    replyMessage(replyToken, getWeekendPlanText());
  }
}


function recordExpense(category, amount, memo, date) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_LOG);
    const now = date ? new Date(date) : new Date();
    sheet.appendRow([
      Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd'),
      Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm'),
      category,
      amount,
      memo || ''
    ]);
  } catch (err) {
    console.log('recordExpense エラー: ' + err.toString());
    throw err;
  }
}

function recordIncome(category, amount, memo, date) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_INCOME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_INCOME);
      sheet.appendRow(['日付', '時刻', 'カテゴリ', '金額', 'メモ']);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    }
    const now = date ? new Date(date) : new Date();
    sheet.appendRow([
      Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd'),
      Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm'),
      category,
      amount,
      memo || ''
    ]);
  } catch (err) {
    console.log('recordIncome エラー: ' + err.toString());
    throw err;
  }
}

function makeBar(ratio) {
  const filled = Math.round(ratio * 10);
  return '█'.repeat(filled) + '░'.repeat(10 - filled);
}

// 今月の完了済み全週の累積超過額を返す（節約は超過を相殺）
function getPrevWeekOverage() {
  const jst = getJstDate(new Date());
  const y = jst.getFullYear();
  const m = jst.getMonth();
  const daysInMonth  = new Date(y, m + 1, 0).getDate();
  const monthStart   = clearTime(new Date(y, m, 1));
  const monthEnd     = clearTime(new Date(y, m + 1, 0));
  const monthBudget  = getMonthlyTotalBudget();
  if (monthBudget <= 0) return 0;

  const weekStart = getWeekStart(jst);

  // 月初を含む週の月曜日を求める
  const firstMon = new Date(monthStart);
  const dow = firstMon.getDay();
  firstMon.setDate(firstMon.getDate() - (dow === 0 ? 6 : dow - 1));

  let carryForward = 0;
  let ws = new Date(firstMon);

  while (ws < weekStart) {           // 今週より前の完了済み週のみ
    const we = new Date(ws);
    we.setDate(we.getDate() + 6);

    // 今月内に収まる有効期間を計算
    const effStart = clearTime(new Date(Math.max(ws.getTime(), monthStart.getTime())));
    const effEnd   = clearTime(new Date(Math.min(we.getTime(), monthEnd.getTime())));

    if (effStart <= effEnd) {
      const days       = Math.round((effEnd - effStart) / 86400000) + 1;
      const weekBudget = Math.round(monthBudget * days / daysInMonth);
      const spent      = sumVariableExpenses(effStart, effEnd);
      carryForward    += (spent - weekBudget);   // 節約はマイナスとして相殺
    }

    ws.setDate(ws.getDate() + 7);
  }

  return carryForward > 0 ? Math.round(carryForward) : 0;
}


function getBalanceReport() {
  try {
    const jst = getJstDate(new Date());
    const weekBudgetBase = getWeeklyBudget();
    const prevOverage = getPrevWeekOverage();
    const weekBudget = Math.max(0, weekBudgetBase - prevOverage);
    const weekStart = getWeekStart(jst);
    const weekTotal = sumVariableExpenses(weekStart, jst);
    const dayOfWeek = jst.getDay() === 0 ? 7 : jst.getDay();
    const todayBalance = Math.floor(weekBudget / 7 * dayOfWeek) - weekTotal;
    const monthStart = new Date(jst.getFullYear(), jst.getMonth(), 1);
    const monthTotal = sumVariableExpenses(monthStart, jst);
    const monthBudget = getMonthlyTotalBudget();

    const weekRemaining = weekBudget - weekTotal;
    const monthRemaining = monthBudget - monthTotal;

    const weekRatio = weekBudget > 0 ? Math.min(weekTotal / weekBudget, 1) : 1;
    const monthRatio = Math.min(monthTotal / monthBudget, 1);
    const dayNames = ['日', '月', '火', '水', '木', '金', '土'];
    const dateStr = `${jst.getMonth() + 1}/${jst.getDate()}(${dayNames[jst.getDay()]})`;

    return (
      `💰 ${dateStr} 残高チェック\n\n` +
      (prevOverage > 0 ? `⚠️ 前週までの超過 -¥${prevOverage.toLocaleString()} 差し引き済み\n\n` : '') +
      `🐟 今日 → ¥${todayBalance.toLocaleString()} 使える\n\n` +
      `📅 今週｜¥${weekTotal.toLocaleString()} / ¥${weekBudget.toLocaleString()}\n` +
      `    ${makeBar(weekRatio)} ${Math.round(weekRatio * 100)}%\n` +
      `    残り ¥${weekRemaining.toLocaleString()}\n\n` +
      `📆 今月｜¥${monthTotal.toLocaleString()} / ¥${monthBudget.toLocaleString()}\n` +
      `    ${makeBar(monthRatio)} ${Math.round(monthRatio * 100)}%\n` +
      `    残り ¥${monthRemaining.toLocaleString()}（変動費のみ）`
    );
  } catch (err) {
    console.log('getBalanceReport エラー: ' + err.toString());
    return 'レポートの取得に失敗しました。';
  }
}

function getConfig() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_CONFIG);
  const data = sheet.getDataRange().getValues();
  const config = {};
  data.forEach(row => {
    if (row[0]) config[row[0]] = Number(row[1]);
  });
  return config;
}

function getMonthlyReport() {
  try {
    const jst = getJstDate(new Date());
    const monthStart = new Date(jst.getFullYear(), jst.getMonth(), 1);
    const monthEnd = new Date(jst.getFullYear(), jst.getMonth() + 1, 0);
    const categories = getCategoryNames();
    const incomeCategories = getIncomeCategoryNames();
    const budgets = getCategoryBudgets();
    const config = getConfig();

    // 収入集計
    const incomeData = getMonthlyIncomeData();
    const incomeByCategory = {};
    incomeCategories.forEach(cat => { incomeByCategory[cat] = 0; });
    incomeData.forEach(row => {
      if (incomeByCategory[row[2]] !== undefined) incomeByCategory[row[2]] += Number(row[3]) || 0;
    });
    const incomeTotal = Object.values(incomeByCategory).reduce((s, v) => s + v, 0);
    const income = incomeTotal > 0 ? incomeTotal : (config['手取り'] || 250000);

    let report = `📊 今月の集計（${jst.getFullYear()}年${jst.getMonth() + 1}月）\n─────────────────\n`;

    report += `【収入】\n`;
    if (incomeTotal > 0) {
      incomeCategories.forEach(cat => {
        if (incomeByCategory[cat] > 0) {
          report += `💰 ${cat}：${incomeByCategory[cat].toLocaleString()}円\n`;
        }
      });
      report += `小計：${incomeTotal.toLocaleString()}円\n\n`;
    } else {
      report += `（未記録 → 手取り設定値：${(config['手取り'] || 250000).toLocaleString()}円を使用）\n\n`;
    }

    let variableSpent = 0;
    let variableBudget = 0;
    report += `【変動費】\n`;
    categories.filter(cat => !FIXED_CATEGORIES.includes(cat)).forEach(cat => {
      const spent = sumExpensesByCategory(monthStart, monthEnd, cat);
      const budget = budgets[cat] || 0;
      report += `${spent <= budget ? '🟢' : '🔴'} ${cat}：${spent.toLocaleString()}円 / ${budget.toLocaleString()}円\n`;
      variableSpent += spent;
      variableBudget += budget;
    });
    report += `小計：${variableSpent.toLocaleString()}円 / ${variableBudget.toLocaleString()}円\n`;

    report += `\n【固定費】\n`;
    let fixedSpent = 0;
    FIXED_CATEGORIES.forEach(cat => {
      const spent = sumExpensesByCategory(monthStart, monthEnd, cat);
      const budget = budgets[cat] || 0;
      report += `${spent <= budget ? '🟢' : '🔴'} ${cat}：${spent.toLocaleString()}円 / ${budget.toLocaleString()}円\n`;
      fixedSpent += spent;
    });

    const savings = income - variableSpent - fixedSpent;
    report += `─────────────────\n`;
    report += `\n💰 収支サマリー\n収入：${income.toLocaleString()}円\n固定費：-${fixedSpent.toLocaleString()}円\n変動費：-${variableSpent.toLocaleString()}円\n貯蓄額：${savings.toLocaleString()}円`;
    return report;
  } catch (err) {
    console.log('getMonthlyReport エラー: ' + err.toString());
    return '月次レポートの取得に失敗しました。';
  }
}

function getMonthlyDetail() {
  try {
    const jst = getJstDate(new Date());
    const monthStart = clearTime(new Date(jst.getFullYear(), jst.getMonth(), 1));
    const monthEnd = clearTime(new Date(jst.getFullYear(), jst.getMonth() + 1, 0));

    const allData = getLogData()
      .filter(row => {
        const d = clearTime(new Date(row[0]));
        return d >= monthStart && d <= monthEnd;
      })
      .sort((a, b) => new Date(a[0]) - new Date(b[0]));

    if (allData.length === 0) {
      return `📋 ${jst.getMonth() + 1}月の明細\n\nまだ記録がありません。`;
    }

    const variableData = allData.filter(row => !FIXED_CATEGORIES.includes(row[2]));
    const fixedData = allData.filter(row => FIXED_CATEGORIES.includes(row[2]));

    const buildSection = (rows) => {
      let text = '';
      let total = 0;
      rows.forEach(row => {
        const d = new Date(row[0]);
        const dateStr = `${d.getMonth() + 1}/${d.getDate()}`;
        const amount = Number(row[3]) || 0;
        const memo = row[4] ? ` ${row[4]}` : '';
        text += `${dateStr} ${row[2]} ${amount.toLocaleString()}円${memo}\n`;
        total += amount;
      });
      return { text, total };
    };

    const variable = buildSection(variableData);
    const fixed = buildSection(fixedData);
    const grandTotal = variable.total + fixed.total;

    let report = `📋 ${jst.getMonth() + 1}月の明細（${allData.length}件）\n`;
    report += `─────────────────\n`;
    report += `【変動費】\n${variable.text}`;
    report += `小計：${variable.total.toLocaleString()}円\n`;
    report += `\n【固定費】\n${fixed.text}`;
    report += `小計：${fixed.total.toLocaleString()}円\n`;
    report += `─────────────────\n合計：${grandTotal.toLocaleString()}円`;

    if (report.length > 4800) {
      report = report.substring(0, 4800) + '\n…（省略）';
    }

    return report;
  } catch (err) {
    console.log('getMonthlyDetail エラー: ' + err.toString());
    return '明細の取得に失敗しました。';
  }
}

function sendMorningReport() {
  try {
    const jst = getJstDate(new Date());
    const weekBudgetBase = getWeeklyBudget();
    const prevOverage = getPrevWeekOverage();
    const weekBudget = Math.max(0, weekBudgetBase - prevOverage);
    const yesterday = new Date(jst);
    yesterday.setDate(yesterday.getDate() - 1);
    const yesterdayTotal = sumVariableExpenses(yesterday, yesterday);
    const weekTotal = sumVariableExpenses(getWeekStart(jst), jst);
    const dayOfWeek = jst.getDay() === 0 ? 7 : jst.getDay();
    const todayBalance = Math.floor(weekBudget / 7 * dayOfWeek) - weekTotal;
    const message =
      `🌅 おはようございます！今日の家計チェックです\n─────────────────\n` +
      `📅 ${formatDate(jst)}（${getDayName(jst.getDay())}）\n\n` +
      `昨日の支出：¥${yesterdayTotal.toLocaleString()}円\n` +
      `今週の累計：¥${weekTotal.toLocaleString()} / ¥${weekBudget.toLocaleString()}円\n` +
      (prevOverage > 0 ? `⚠️ 前週までの超過 -¥${prevOverage.toLocaleString()} 差し引き済み\n` : '') +
      `\n` +
      `🟢 今日使える金額：${todayBalance.toLocaleString()}円\n─────────────────\n今日も良い1日を！`;
    pushMessage(MY_USER_ID, message);
  } catch (err) {
    console.log('sendMorningReport エラー: ' + err.toString());
  }
}

function sendMonthlyReport() {
  try {
    pushMessage(MY_USER_ID, `📆 今月の家計まとめをお届けします！\n\n${getMonthlyReport()}\n\n来月もコツコツいきましょう💪`);
  } catch (err) {
    console.log('sendMonthlyReport エラー: ' + err.toString());
  }
}

function getCommandList() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_COMMANDS);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 3).getValues()
    .filter(row => row[0] !== '');
}

function getHelpMessage() {
  const commands = getCommandList();
  let text = `📋 コマンド一覧\n─────────────────\n`;
  commands.forEach(row => {
    text += `▶ ${row[0]}\n   ${row[1]}\n`;
  });
  text += `─────────────────\n💡 カテゴリ名でも記録できます\n例：「食費 コンビニ 400」`;
  text += getQuickLogHelp();
  return text;
}

function replyMessage(replyToken, text, quickReplies) {
  callLineApi('https://api.line.me/v2/bot/message/reply', { replyToken, messages: [buildTextMessage(text, quickReplies)] });
}

function pushMessage(userId, text, quickReplies) {
  callLineApi('https://api.line.me/v2/bot/message/push', { to: userId, messages: [buildTextMessage(text, quickReplies)] });
}

function buildTextMessage(text, quickReplies) {
  const msg = { type: 'text', text };
  if (quickReplies && quickReplies.length > 0) {
    msg.quickReply = { items: quickReplies.map(qr => ({ type: 'action', action: qr })) };
  }
  return msg;
}

function sumExpenses(startDate, endDate) {
  const start = clearTime(startDate);
  const end = clearTime(endDate);
  return getLogData().reduce((total, row) => {
    const d = clearTime(new Date(row[0]));
    return (d >= start && d <= end) ? total + (Number(row[3]) || 0) : total;
  }, 0);
}

function sumVariableExpenses(startDate, endDate) {
  const start = clearTime(startDate);
  const end = clearTime(endDate);
  const variableCats = getCategoryNames().filter(cat => !FIXED_CATEGORIES.includes(cat));
  return getLogData().reduce((total, row) => {
    const d = clearTime(new Date(row[0]));
    return (d >= start && d <= end && variableCats.includes(row[2]))
      ? total + (Number(row[3]) || 0)
      : total;
  }, 0);
}

function sumExpensesByCategory(startDate, endDate, category) {
  const start = clearTime(startDate);
  const end = clearTime(endDate);
  return getLogData().reduce((total, row) => {
    const d = clearTime(new Date(row[0]));
    return (d >= start && d <= end && row[2] === category) ? total + (Number(row[3]) || 0) : total;
  }, 0);
}

function getLogData() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_LOG);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 5).getValues();
}

function getMonthlyIncomeData() {
  const jst = getJstDate(new Date());
  const monthStart = clearTime(new Date(jst.getFullYear(), jst.getMonth(), 1));
  const monthEnd = clearTime(new Date(jst.getFullYear(), jst.getMonth() + 1, 0));
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_INCOME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues()
    .filter(row => {
      const d = clearTime(new Date(row[0]));
      return d >= monthStart && d <= monthEnd;
    });
}

function getMonthlyIncomeTotal() {
  return getMonthlyIncomeData().reduce((sum, row) => sum + (Number(row[3]) || 0), 0);
}

function getCategoryNames() {
  const EXCLUDE = ['週予算', '手取り'];
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_CONFIG);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 3).getValues()
    .filter(row => row[0] !== '' && !EXCLUDE.includes(row[0]) && row[2] !== '収入')
    .map(row => row[0]);
}

function getIncomeCategoryNames() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_CONFIG);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 3).getValues()
    .filter(row => row[2] === '収入' && row[0] !== '')
    .map(row => row[0]);
}

function getCategoryBudgets() {
  const EXCLUDE = ['週予算', '手取り'];
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_CONFIG);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  return sheet.getRange(2, 1, lastRow - 1, 3).getValues().reduce((obj, row) => {
    if (row[0] && !EXCLUDE.includes(row[0]) && row[2] !== '収入') obj[row[0]] = Number(row[1]) || 0;
    return obj;
  }, {});
}

function getWeeklyBudget() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_CONFIG);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 17000;
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const row = data.find(r => r[0] === '週予算');
  return row ? (Number(row[1]) || 17000) : 17000;
}

function getMonthlyTotalBudget() {
  const budgets = getCategoryBudgets();
  return Object.entries(budgets)
    .filter(([cat]) => !FIXED_CATEGORIES.includes(cat))
    .reduce((sum, [, v]) => sum + v, 0);
}

function setupSpreadsheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let logSheet = ss.getSheetByName(SHEET_LOG);
  if (!logSheet) logSheet = ss.insertSheet(SHEET_LOG);
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(['日付', '時刻', 'カテゴリ', '金額', 'メモ']);
    logSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  }
  let configSheet = ss.getSheetByName(SHEET_CONFIG);
  if (!configSheet) configSheet = ss.insertSheet(SHEET_CONFIG);
  if (configSheet.getLastRow() === 0) {
    configSheet.appendRow(['カテゴリ', '月予算（円）', '備考']);
    configSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    [
      ['食費', 22000, ''],
      ['交際費', 35000, ''],
      ['交通費', 10000, ''],
      ['娯楽', 8000, ''],
      ['日用品', 5000, ''],
      ['固定費', 90000, '家賃・奨学金・サブスク・携帯'],
      ['週予算', 20000, '週の変動費予算（残高確認で使用）'],
      ['手取り', 250000, '給与'],
      ['給料', 0, '収入'],
      ['副業', 0, '収入'],
      ['その他収入', 0, '収入'],
    ].forEach(row => configSheet.appendRow(row));
  }
  let incomeSheet = ss.getSheetByName(SHEET_INCOME);
  if (!incomeSheet) {
    incomeSheet = ss.insertSheet(SHEET_INCOME);
    incomeSheet.appendRow(['日付', '時刻', 'カテゴリ', '金額', 'メモ']);
    incomeSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  }
  console.log('セットアップ完了');
}

function setupCommandSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_COMMANDS);
  if (!sheet) sheet = ss.insertSheet(SHEET_COMMANDS);
  sheet.clearContents();
  sheet.appendRow(['コマンド名', '説明', '固定返答']);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  [
    ['残高確認', '今日使える金額・今週今月の進捗', ''],
    ['今月の集計', 'カテゴリ別の支出と予算', ''],
    ['今月の明細', '全支出の一覧（変動費・固定費別）', ''],
    ['記録する', 'カテゴリを選んで支出を記録', ''],
    ['ヘルプ', 'コマンド一覧を表示', ''],
  ].forEach(row => sheet.appendRow(row));
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 200);
  console.log('コマンドシート作成完了');
}

function getJstDate(date) {
  return new Date(date.getTime() + date.getTimezoneOffset() * 60000 + 9 * 3600000);
}

function clearTime(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function getWeekStart(date) {
  const d = new Date(date);
  const day = d.getDay();
  d.setDate(d.getDate() + (day === 0 ? -6 : 1 - day));
  return clearTime(d);
}

function formatDate(date) {
  return `${date.getMonth() + 1}月${date.getDate()}日`;
}

function getDayName(dayIndex) {
  return ['日', '月', '火', '水', '木', '金', '土'][dayIndex] + '曜日';
}

function parseQuery(queryString) {
  return queryString.split('&').reduce((obj, pair) => {
    const [key, value] = pair.split('=');
    obj[decodeURIComponent(key)] = decodeURIComponent(value || '');
    return obj;
  }, {});
}

function testLineApi() {
  console.log('=== LINE API テスト開始 ===');
  console.log('ACCESS_TOKEN の最初の10文字: ' + LINE_CHANNEL_ACCESS_TOKEN.substring(0, 10));
  try {
    const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/info', {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN },
      muteHttpExceptions: true
    });
    console.log('Bot Info レスポンス: ' + res.getContentText());
    console.log('ステータスコード: ' + res.getResponseCode());
  } catch(e) {
    console.log('Bot Info エラー: ' + e.toString());
  }
}

function testPush() {
  pushMessage(MY_USER_ID, getBalanceReport());
}

function setupUsageSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('使い方');
  if (!sheet) sheet = ss.insertSheet('使い方');
  sheet.clearContents();

  const data = [
    ['LINE家計簿 使い方'],
    [''],
    ['■ 支出の記録'],
    ['方法①', '一行入力', '「食費 コンビニ 400」のように送る'],
    ['方法②', 'カテゴリ選択', 'カテゴリ名を送る → 金額を送る'],
    ['方法③', 'LIFFフォーム', 'メニューの「記録する」から入力'],
    [''],
    ['■ 収入の記録'],
    ['方法①', '一行入力', '「給料 250000」のように送る'],
    ['方法②', 'カテゴリ選択', '「給料」「副業」などを送る → 金額を送る'],
    [''],
    ['■ 確認コマンド'],
    ['残高確認', '今日使える金額・今週・今月の進捗を表示'],
    ['今月の集計', 'カテゴリ別の収入・支出と予算を表示'],
    ['今月の明細', '全支出を変動費・固定費別に一覧表示'],
    ['ヘルプ', 'コマンド一覧を表示'],
    [''],
    ['■ 支出カテゴリ一覧（設定シートで変更可）'],
    ['食費', '交際費', '交通費', '娯楽', '日用品', '固定費'],
    [''],
    ['■ 収入カテゴリ（設定シートのC列に「収入」と記入）'],
    ['給料', '副業', 'その他収入'],
    [''],
    ['■ 予算の変更方法'],
    ['月予算', '設定シートのB列の数字を変える'],
    ['週予算', '設定シートの「週予算」行のB列を変える'],
    [''],
    ['■ 自動通知'],
    ['朝8時', '昨日の支出・今日使える金額をLINEに送信'],
    ['月末', '今月の集計をLINEに送信'],
  ];

  sheet.getRange(1, 1, data.length, 6).setValues(
    data.map(row => {
      while (row.length < 6) row.push('');
      return row;
    })
  );

  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 300);

  console.log('使い方シート作成完了');
}

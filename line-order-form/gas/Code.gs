/**
 * LINE注文フォーム システム
 * ------------------------------------------------------------
 * Googleスプレッドシートをデータベースとして使用する注文管理ウェブアプリ。
 *
 * シート構成:
 *   - 商品マスター : 商品ID / 商品名 / 価格(元) / 表示順 / 有効
 *   - 注文一覧     : 注文明細(1商品1行、同一注文は注文IDで紐付け)
 *   - 設定         : 振込先口座・管理者キー・通知メールなど
 *
 * 使い方は SETUP.md を参照してください。
 */

// ===== 定数 =====
var SHEET_PRODUCTS = '商品マスター';
var SHEET_ORDERS   = '注文一覧';
var SHEET_SETTINGS = '設定';
var TIMEZONE       = 'Asia/Taipei';

var STATUSES = ['未対応', '入金確認済', '梱包済', '受渡済', 'キャンセル'];

var ORDER_HEADERS = [
  '注文ID', '注文日時', 'オープンチャット名', 'LINEアカウント名',
  '商品名', '数量', '単価(元)', '小計(元)', '注文合計(元)',
  'ステータス', '入力者', '備考'
];

// ===== エントリーポイント =====

/**
 * ウェブアプリのGETリクエスト処理。
 *   通常アクセス          → お客様用注文フォーム
 *   ?page=admin&key=XXX  → 管理者ページ(キー認証)
 */
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) || '';

  if (page === 'admin') {
    var key = (e.parameter.key || '');
    if (!isValidAdminKey_(key)) {
      return HtmlService.createHtmlOutput(
        '<meta name="viewport" content="width=device-width, initial-scale=1">' +
        '<div style="font-family:sans-serif;padding:2em;text-align:center;">' +
        '<h2>アクセスできません</h2>' +
        '<p>管理者キーが正しくありません。<br>URLの key パラメータをご確認ください。</p></div>'
      ).setTitle('管理者ページ');
    }
    var tmpl = HtmlService.createTemplateFromFile('Admin');
    tmpl.adminKey = key;
    return tmpl.evaluate()
      .setTitle('注文管理(管理者)')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  var isAdminEntry = page === 'staff' && isValidAdminKey_(e.parameter.key || '');
  var formTmpl = HtmlService.createTemplateFromFile('Form');
  formTmpl.isAdminEntry = isAdminEntry;
  return formTmpl.evaluate()
    .setTitle('ご注文フォーム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ===== 公開API(google.script.run から呼び出し) =====

/**
 * フォーム初期表示用データ: 有効な商品一覧と振込先情報。
 */
function getFormInit() {
  return {
    products: getActiveProducts_(),
    payment: getPaymentInfo_()
  };
}

/**
 * 注文を登録する。
 * @param {Object} data {openChatName, lineName, note, enteredBy, items:[{productId, name, qty}]}
 * @return {Object} {orderId, orderedAt, items, total, hasCustomItems, payment}
 */
function submitOrder(data) {
  if (!data || !data.openChatName || !String(data.openChatName).trim()) {
    throw new Error('オープンチャット名を入力してください。');
  }
  var rawItems = (data.items || []).filter(function (it) {
    return it && Number(it.qty) > 0 && (it.productId || (it.name && String(it.name).trim()));
  });
  if (rawItems.length === 0) {
    throw new Error('商品を1つ以上選択してください。');
  }

  // 価格・商品名は必ずサーバー側でマスターから引き直す(改ざん防止)
  var master = {};
  getActiveProducts_().forEach(function (p) { master[p.id] = p; });

  var items = rawItems.map(function (it) {
    var qty = Math.floor(Number(it.qty));
    if (!(qty > 0) || qty > 9999) throw new Error('数量が正しくありません。');
    var m = it.productId ? master[String(it.productId)] : null;
    if (m) {
      return { name: m.name, qty: qty, price: m.price, subtotal: m.price === '' ? '' : m.price * qty, custom: false };
    }
    return { name: String(it.name).trim(), qty: qty, price: '', subtotal: '', custom: true };
  });

  var total = items.reduce(function (sum, it) {
    return sum + (it.subtotal === '' ? 0 : it.subtotal);
  }, 0);
  var hasCustomItems = items.some(function (it) { return it.custom; });

  var lock = LockService.getScriptLock();
  lock.waitLock(20 * 1000);
  var orderId, orderedAt;
  try {
    var sheet = getSheet_(SHEET_ORDERS);
    orderId = getNextOrderId_(sheet);
    orderedAt = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy/MM/dd HH:mm');

    var enteredBy = data.enteredBy === '管理者' ? '管理者' : 'お客様';
    var rows = items.map(function (it, i) {
      return [
        orderId,
        orderedAt,
        String(data.openChatName).trim(),
        String(data.lineName || '').trim(),
        it.name,
        it.qty,
        it.price,
        it.subtotal,
        i === 0 ? total : '',   // 注文合計は先頭行のみ(SUMの二重計上防止)
        STATUSES[0],
        enteredBy,
        i === 0 ? String(data.note || '').trim() : ''
      ];
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, ORDER_HEADERS.length).setValues(rows);
  } finally {
    lock.releaseLock();
  }

  notifyNewOrder_(orderId, orderedAt, data, items, total, hasCustomItems);

  return {
    orderId: orderId,
    orderedAt: orderedAt,
    items: items,
    total: total,
    hasCustomItems: hasCustomItems,
    payment: getPaymentInfo_()
  };
}

/**
 * 管理者ページ用: 全注文データを注文ID単位にまとめて返す。
 */
function getAdminData(key) {
  requireAdmin_(key);
  var sheet = getSheet_(SHEET_ORDERS);
  var lastRow = sheet.getLastRow();
  var orders = {};
  var orderIdsInOrder = [];

  if (lastRow >= 2) {
    var values = sheet.getRange(2, 1, lastRow - 1, ORDER_HEADERS.length).getValues();
    values.forEach(function (r) {
      var id = String(r[0]);
      if (!id) return;
      if (!orders[id]) {
        orders[id] = {
          orderId: id,
          orderedAt: r[1] instanceof Date
            ? Utilities.formatDate(r[1], TIMEZONE, 'yyyy/MM/dd HH:mm') : String(r[1]),
          openChatName: String(r[2]),
          lineName: String(r[3]),
          status: String(r[9]) || STATUSES[0],
          enteredBy: String(r[10]),
          note: String(r[11] || ''),
          total: 0,
          hasCustomItems: false,
          items: []
        };
        orderIdsInOrder.push(id);
      }
      var price = r[6], subtotal = r[7];
      var custom = (price === '' || price === null);
      orders[id].items.push({
        name: String(r[4]),
        qty: Number(r[5]) || 0,
        price: custom ? '' : Number(price),
        subtotal: (subtotal === '' || subtotal === null) ? '' : Number(subtotal)
      });
      if (custom) orders[id].hasCustomItems = true;
      if (subtotal !== '' && subtotal !== null) orders[id].total += Number(subtotal);
      if (!orders[id].note && r[11]) orders[id].note = String(r[11]);
    });
  }

  return {
    statuses: STATUSES,
    orders: orderIdsInOrder.map(function (id) { return orders[id]; }).reverse() // 新しい順
  };
}

/**
 * 管理者ページ用: 注文のステータスを更新する(同一注文IDの全行)。
 */
function updateOrderStatus(key, orderId, status) {
  requireAdmin_(key);
  if (STATUSES.indexOf(status) === -1) throw new Error('不正なステータスです。');

  var lock = LockService.getScriptLock();
  lock.waitLock(20 * 1000);
  try {
    var sheet = getSheet_(SHEET_ORDERS);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { updated: 0 };
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var updated = 0;
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(orderId)) {
        sheet.getRange(i + 2, 10).setValue(status); // 10列目 = ステータス
        updated++;
      }
    }
    return { updated: updated };
  } finally {
    lock.releaseLock();
  }
}

// ===== 内部ヘルパー =====

function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('シート「' + name + '」が見つかりません。メニューの「初期セットアップ」を実行してください。');
  }
  return sheet;
}

/** 商品マスターから有効な商品を表示順で返す。 */
function getActiveProducts_() {
  var sheet = getSheet_(SHEET_PRODUCTS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var products = [];
  values.forEach(function (r) {
    var id = String(r[0]).trim();
    var name = String(r[1]).trim();
    if (!id || !name) return;
    var active = String(r[4]).trim();
    if (active === '無効' || active === 'FALSE' || active === 'false' || r[4] === false) return;
    var price = (r[2] === '' || r[2] === null || isNaN(Number(r[2]))) ? '' : Number(r[2]);
    products.push({
      id: id,
      name: name,
      price: price,
      sortOrder: (r[3] === '' || isNaN(Number(r[3]))) ? 9999 : Number(r[3])
    });
  });
  products.sort(function (a, b) { return a.sortOrder - b.sortOrder; });
  return products;
}

/** 設定シートをキー・バリューのオブジェクトとして読み込む。 */
function getSettings_() {
  var sheet = getSheet_(SHEET_SETTINGS);
  var lastRow = sheet.getLastRow();
  var settings = {};
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 2).getValues().forEach(function (r) {
      var k = String(r[0]).trim();
      if (k) settings[k] = String(r[1]).trim();
    });
  }
  return settings;
}

function getPaymentInfo_() {
  var s = getSettings_();
  return {
    bankName:      s['振込先_銀行名'] || '台湾 中華郵政(郵便局)',
    bankCode:      s['振込先_郵局代号'] || '700',
    accountNumber: s['振込先_口座番号'] || '',
    accountName:   s['振込先_口座名義'] || '',
    paymentNote:   s['振込メモ'] || ''
  };
}

function isValidAdminKey_(key) {
  var s = getSettings_();
  var adminKey = s['管理者キー'] || '';
  return adminKey !== '' && key === adminKey;
}

function requireAdmin_(key) {
  if (!isValidAdminKey_(key)) throw new Error('管理者キーが正しくありません。');
}

/** 注文IDを採番する。形式: yyyyMMdd-連番3桁 (例: 20260811-003) */
function getNextOrderId_(sheet) {
  var prefix = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMdd');
  var lastRow = sheet.getLastRow();
  var seen = {};
  var count = 0;
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(function (r) {
      var id = String(r[0]);
      if (id.indexOf(prefix + '-') === 0 && !seen[id]) {
        seen[id] = true;
        count++;
      }
    });
  }
  return prefix + '-' + ('00' + (count + 1)).slice(-3);
}

/** 新規注文をメールで管理者へ通知する(設定シートに通知メールがある場合のみ)。 */
function notifyNewOrder_(orderId, orderedAt, data, items, total, hasCustomItems) {
  try {
    var email = getSettings_()['通知メール'] || '';
    if (!email) return;
    var lines = [
      '新しい注文が入りました。',
      '',
      '注文ID: ' + orderId,
      '日時: ' + orderedAt,
      'オープンチャット名: ' + data.openChatName,
      'LINEアカウント名: ' + (data.lineName || '(未入力)'),
      '入力者: ' + (data.enteredBy === '管理者' ? '管理者' : 'お客様'),
      '',
      '--- 注文内容 ---'
    ];
    items.forEach(function (it) {
      lines.push(it.name + ' × ' + it.qty +
        (it.price === '' ? '(価格未定)' : '(' + it.price + '元 / 小計 ' + it.subtotal + '元)'));
    });
    lines.push('');
    lines.push('合計: ' + total + '元' + (hasCustomItems ? '(価格未定の商品を除く)' : ''));
    if (data.note) lines.push('備考: ' + data.note);
    MailApp.sendEmail(email, '【注文】' + orderId + ' ' + data.openChatName, lines.join('\n'));
  } catch (err) {
    // 通知失敗で注文自体を失敗させない
    console.error('通知メール送信エラー: ' + err);
  }
}

// ===== 初期セットアップ =====

/** スプレッドシートを開いたときにメニューを追加する。 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('注文フォーム')
    .addItem('初期セットアップ(シート作成)', 'setupSheets')
    .addToUi();
}

/**
 * 必要なシートとヘッダー・サンプルデータを作成する。
 * 既存シートがある場合は変更しない(安全)。
 */
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(SHEET_PRODUCTS)) {
    var p = ss.insertSheet(SHEET_PRODUCTS);
    p.getRange(1, 1, 1, 5).setValues([['商品ID', '商品名', '価格(元)', '表示順', '有効']])
      .setFontWeight('bold').setBackground('#d9ead3');
    p.getRange(2, 1, 2, 5).setValues([
      ['P001', 'サンプル商品A', 120, 1, '有効'],
      ['P002', 'サンプル商品B', 250, 2, '有効']
    ]);
    p.setFrozenRows(1);
  }

  if (!ss.getSheetByName(SHEET_ORDERS)) {
    var o = ss.insertSheet(SHEET_ORDERS);
    o.getRange(1, 1, 1, ORDER_HEADERS.length).setValues([ORDER_HEADERS])
      .setFontWeight('bold').setBackground('#cfe2f3');
    o.setFrozenRows(1);
    // ステータス列に入力規則(プルダウン)を設定
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build();
    o.getRange(2, 10, o.getMaxRows() - 1, 1).setDataValidation(rule);
  }

  if (!ss.getSheetByName(SHEET_SETTINGS)) {
    var s = ss.insertSheet(SHEET_SETTINGS);
    s.getRange(1, 1, 1, 3).setValues([['設定キー', '値', '説明']])
      .setFontWeight('bold').setBackground('#fce5cd');
    s.getRange(2, 1, 7, 3).setValues([
      ['振込先_銀行名', '台湾 中華郵政(郵便局)', '振込先の銀行名(表示用)'],
      ['振込先_郵局代号', '700', '中華郵政の局番コード'],
      ['振込先_口座番号', '', '例: 0000000-0000000 ← 必ずご自身の口座番号に変更'],
      ['振込先_口座名義', '', '口座名義人のお名前'],
      ['振込メモ', '振込後、オープンチャットでお知らせください。', '注文完了画面に表示される案内文'],
      ['管理者キー', generateKey_(), '管理者ページURLの key= に使う秘密の文字列。変更可'],
      ['通知メール', '', '新規注文をメール通知したい場合にアドレスを入力(空欄なら通知なし)']
    ]);
    s.setFrozenRows(1);
    s.setColumnWidth(2, 260);
    s.setColumnWidth(3, 380);
  }

  SpreadsheetApp.getUi().alert(
    'セットアップ完了!\n\n' +
    '1.「設定」シートで振込先口座を入力してください。\n' +
    '2.「商品マスター」シートに商品を登録してください。\n' +
    '3. デプロイしてURLを取得してください(SETUP.md参照)。'
  );
}

/** ランダムな管理者キーを生成する。 */
function generateKey_() {
  var chars = 'abcdefghijkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var key = '';
  for (var i = 0; i < 16; i++) {
    key += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return key;
}

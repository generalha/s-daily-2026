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
var SHEET_PRODUCTS  = '商品マスター';
var SHEET_ORDERS    = '注文一覧';
var SHEET_SETTINGS  = '設定';
var SHEET_SHIPPING  = '送料マスター';
var SHEET_CUSTOMERS = '顧客マスター';
var TIMEZONE       = 'Asia/Taipei';

// フロー: 承認待ち →(管理者承認・在庫引落)→ 注文確定 →(商品代入金)→ 入金確認済
//   →(計量・送料確定)→ 梱包済 →(送料入金)→ 送料入金確認済
//   →(発送ラベル作成・発送)→ 発送済 →(到着確認)→ 受渡済
var STATUSES = ['承認待ち', '注文確定', '入金確認済', '梱包済', '送料入金確認済', '発送済', '受渡済', 'キャンセル'];
// 在庫引落済みのステータス(これらに入る時に在庫を減らし、抜ける時に在庫を戻す)
var STOCK_TAKEN_STATUSES = ['注文確定', '入金確認済', '梱包済', '送料入金確認済', '発送済', '受渡済'];

var ORDER_HEADERS = [
  '注文ID', '注文日時', 'オープンチャット名', 'LINEアカウント名',
  '商品名', '数量', '単価(元)', '小計(元)', '注文合計(元)',
  'ステータス', '入力者', '備考',
  '重量(g)', '配送方法', '国際送料(元)',
  'お届け先氏名', 'お届け先住所', 'お届け先電話', '追跡番号', '発送日'
];
// 列番号(1始まり)
var COL_STATUS = 10, COL_WEIGHT = 13, COL_METHOD = 14, COL_SHIPFEE = 15,
    COL_RECV_NAME = 16, COL_RECV_ADDR = 17, COL_RECV_PHONE = 18,
    COL_TRACKING = 19, COL_SHIPDATE = 20;

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

  if (page === 'status') {
    var st = HtmlService.createTemplateFromFile('Status');
    st.prefillOrder = String(e.parameter.order || '').replace(/[^0-9\-]/g, '');
    return st.evaluate()
      .setTitle('注文状況の確認')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  var isAdminEntry = page === 'staff' && isValidAdminKey_(e.parameter.key || '');
  var formTmpl = HtmlService.createTemplateFromFile('Form');
  formTmpl.isAdminEntry = isAdminEntry;
  // 顧客No読み込み用。管理者キー検証済みの代理入力モードのみ渡す
  formTmpl.adminKey = isAdminEntry ? String(e.parameter.key || '') : '';
  return formTmpl.evaluate()
    .setTitle('ご注文フォーム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// お客様向け進捗ステップ(キャンセルを除くSTATUSESと1対1対応)
var CUSTOMER_STEPS = [
  { label: '注文受付', desc: '在庫を確認しています。確定のご連絡をお待ちください。' },
  { label: '注文確定', desc: 'ご注文が確定しました。商品代のお振込をお願いします。' },
  { label: '入金確認', desc: '商品代のご入金を確認しました。梱包までお待ちください。' },
  { label: '梱包完了', desc: '梱包が完了しました。国際送料のご案内をお送りしますので、送料のお振込をお願いします。' },
  { label: '送料入金確認', desc: '送料のご入金を確認しました。発送の準備をしています。' },
  { label: '発送済み', desc: '発送しました!お届けまで今しばらくお待ちください。' },
  { label: '受取完了', desc: 'お受け取りが確認されました。ありがとうございました!' }
];

// ===== 公開API(google.script.run から呼び出し) =====

/**
 * フォーム初期表示用データ: 有効な商品一覧と振込先情報。
 */
function getFormInit() {
  return {
    products: getActiveProducts_(),
    payment: getPaymentInfo_(),
    statusUrl: ScriptApp.getService().getUrl() + '?page=status'
  };
}

/**
 * お客様向け: 注文番号+オープンチャット名で注文の進捗を照会する。
 * 両方が一致しないと何も返さない(他人の注文は見られない)。
 */
function getOrderStatus(orderId, openChatName) {
  var id = String(orderId || '').trim();
  var name = String(openChatName || '').trim();
  if (!id || !name) throw new Error('注文番号とオープンチャット名を入力してください。');

  var order = findOrder_(id);
  if (!order || order.openChatName.trim() !== name) {
    throw new Error('注文が見つかりません。注文番号とオープンチャット名(注文時と同じ表記)をご確認ください。');
  }

  var cancelled = order.status === 'キャンセル';
  var statusIndex = STATUSES.indexOf(order.status);
  return {
    orderId: order.orderId,
    orderedAt: order.orderedAt,
    status: order.status,
    statusIndex: cancelled ? -1 : statusIndex,
    steps: CUSTOMER_STEPS,
    cancelled: cancelled,
    items: order.items,
    total: order.total,
    hasCustomItems: order.hasCustomItems,
    shipMethod: order.shipMethod,
    shipFee: order.shipFee,
    weightG: order.weightG,
    trackingNo: order.trackingNo,
    shippedDate: order.shippedDate,
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
      // 在庫の事前チェック(確定は管理者の承認時。同時注文分は承認時に検出される)
      if (m.stock !== '' && qty > m.stock) {
        throw new Error('「' + m.name + '」は残り' + m.stock + '点です。数量を減らしてください。');
      }
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
        i === 0 ? String(data.note || '').trim() : '',
        '', '', '',             // 重量・配送方法・国際送料(梱包時に入力)
        i === 0 ? String(data.recvName || '').trim() : '',
        i === 0 ? String(data.recvAddress || '').trim() : '',
        i === 0 ? String(data.recvPhone || '').trim() : '',
        '', ''                  // 追跡番号・発送日(発送時に入力)
      ];
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, ORDER_HEADERS.length).setValues(rows);
  } finally {
    lock.releaseLock();
  }

  // 顧客マスターへ自動登録・更新(顧客Noを自動採番)
  upsertCustomer_({
    openChatName: data.openChatName,
    lineName: data.lineName,
    recvName: data.recvName,
    recvAddress: data.recvAddress,
    recvPhone: data.recvPhone
  });

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
  return {
    statuses: STATUSES,
    products: getActiveProducts_(),
    shippingRates: getShippingRates_(),
    payment: getPaymentInfo_(),
    statusUrl: ScriptApp.getService().getUrl() + '?page=status',
    customers: readCustomers_(),
    orders: readOrders_().reverse() // 新しい順
  };
}

/**
 * 管理者ページ用: 計量した重量と配送方法から国際送料を確定し、
 * 注文を「梱包済」にする。送料はサーバー側で送料マスターから計算する。
 */
function setPacked(key, orderId, weightG, method) {
  requireAdmin_(key);
  var w = Math.ceil(Number(weightG));
  if (!(w > 0) || w > 100000) throw new Error('重量(g)を正しく入力してください。');
  var fee = calcShippingFee_(method, w);
  if (fee === null) {
    throw new Error('「' + method + '」の料金表に ' + w + 'g に該当する重量帯がありません。送料マスターをご確認ください。');
  }

  var fields = {};
  fields[COL_WEIGHT] = w;
  fields[COL_METHOD] = method;
  fields[COL_SHIPFEE] = fee;
  setOrderFields_(orderId, fields, '梱包済');
  return { orderId: orderId, weightG: w, method: method, fee: fee };
}

/**
 * 管理者ページ用: お届け先・追跡番号を保存する。
 * markAsShipped が true の場合は発送日を記録して「発送済」にする。
 */
function saveShipping(key, orderId, info, markAsShipped) {
  requireAdmin_(key);
  info = info || {};
  var fields = {};
  fields[COL_RECV_NAME]  = String(info.recvName || '').trim();
  fields[COL_RECV_ADDR]  = String(info.recvAddress || '').trim();
  fields[COL_RECV_PHONE] = String(info.recvPhone || '').trim();
  fields[COL_TRACKING]   = String(info.trackingNo || '').trim();

  var newStatus = null;
  if (markAsShipped) {
    fields[COL_SHIPDATE] = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy/MM/dd');
    newStatus = '発送済';
  }
  setOrderFields_(orderId, fields, newStatus);

  // お届け先情報を顧客マスターにも反映(次回の代理入力で使えるように)
  var order = findOrder_(orderId);
  if (order) {
    upsertCustomer_({
      openChatName: order.openChatName,
      lineName: order.lineName,
      recvName: info.recvName,
      recvAddress: info.recvAddress,
      recvPhone: info.recvPhone
    });
  }
  return { orderId: orderId, shipped: !!markAsShipped, shippedDate: fields[COL_SHIPDATE] || '' };
}

/**
 * 管理者ページ用: 注文番号で1件の注文を呼び出す(修正画面用)。
 * 暗号キー(管理者キー)がなければ呼び出せない。
 */
function getOrderById(key, orderId) {
  requireAdmin_(key);
  var order = findOrder_(orderId);
  if (!order) throw new Error('注文番号「' + orderId + '」は見つかりません。');

  // 商品名がマスターに一致する明細には商品IDを付与(修正画面のプルダウン用)
  var byName = {};
  getActiveProducts_().forEach(function (p) { byName[p.name] = p; });
  order.items.forEach(function (it) {
    it.productId = byName[it.name] ? byName[it.name].id : '';
  });
  return order;
}

/**
 * 管理者ページ用: 注文内容を修正する(暗号キー必須)。
 * 注文日時・ステータス・入力者は維持し、明細・お客様情報・備考を差し替える。
 * @param {Object} data {openChatName, lineName, note, items:[{productId, name, qty, price}]}
 */
function updateOrder(key, orderId, data) {
  requireAdmin_(key);
  if (!data || !data.openChatName || !String(data.openChatName).trim()) {
    throw new Error('オープンチャット名を入力してください。');
  }
  var rawItems = (data.items || []).filter(function (it) {
    return it && Number(it.qty) > 0 && (it.productId || (it.name && String(it.name).trim()));
  });
  if (rawItems.length === 0) {
    throw new Error('商品を1つ以上入力してください。(注文の取消はステータスを「キャンセル」にしてください)');
  }

  // マスター商品は現在のマスター価格で再計算、自由入力商品は指定価格を採用
  var master = {};
  getActiveProducts_().forEach(function (p) { master[p.id] = p; });
  var items = rawItems.map(function (it) {
    var qty = Math.floor(Number(it.qty));
    if (!(qty > 0) || qty > 9999) throw new Error('数量が正しくありません。');
    var m = it.productId ? master[String(it.productId)] : null;
    if (m) {
      return { name: m.name, qty: qty, price: m.price, subtotal: m.price === '' ? '' : m.price * qty };
    }
    var price = (it.price === '' || it.price === null || it.price === undefined || isNaN(Number(it.price)))
      ? '' : Number(it.price);
    return { name: String(it.name).trim(), qty: qty, price: price, subtotal: price === '' ? '' : price * qty };
  });
  var total = items.reduce(function (sum, it) {
    return sum + (it.subtotal === '' ? 0 : it.subtotal);
  }, 0);

  var lock = LockService.getScriptLock();
  lock.waitLock(20 * 1000);
  try {
    var sheet = getSheet_(SHEET_ORDERS);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('注文番号「' + orderId + '」は見つかりません。');

    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var rowIdxs = []; // シート上の行番号(1始まり)
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(orderId)) rowIdxs.push(i + 2);
    }
    if (rowIdxs.length === 0) throw new Error('注文番号「' + orderId + '」は見つかりません。');

    // 既存行から維持する情報を取得
    var firstOld = sheet.getRange(rowIdxs[0], 1, 1, ORDER_HEADERS.length).getValues()[0];
    var orderedAt = firstOld[1] instanceof Date
      ? Utilities.formatDate(firstOld[1], TIMEZONE, 'yyyy/MM/dd HH:mm') : String(firstOld[1]);
    var status = String(firstOld[9]) || STATUSES[0];
    var enteredBy = String(firstOld[10]) || 'お客様';

    var editMark = '【' + Utilities.formatDate(new Date(), TIMEZONE, 'M/d HH:mm') + ' 修正】';
    var note = String(data.note || '').trim();
    // 発送関連(重量〜発送日)は既存の値を維持する
    var shipping = firstOld.slice(COL_WEIGHT - 1, COL_SHIPDATE);
    var newRows = items.map(function (it, idx) {
      return [
        orderId, orderedAt,
        String(data.openChatName).trim(), String(data.lineName || '').trim(),
        it.name, it.qty, it.price, it.subtotal,
        idx === 0 ? total : '',
        status, enteredBy,
        idx === 0 ? (editMark + (note ? ' ' + note : '')) : ''
      ].concat(idx === 0 ? shipping : ['', '', '', '', '', '', '', '']);
    });

    // 既存行を下から削除し、元の位置に新しい行を挿入する
    for (var d = rowIdxs.length - 1; d >= 0; d--) {
      sheet.deleteRow(rowIdxs[d]);
    }
    var insertAt = rowIdxs[0];
    sheet.insertRowsBefore(Math.min(insertAt, sheet.getLastRow() + 1), newRows.length);
    sheet.getRange(insertAt, 1, newRows.length, ORDER_HEADERS.length).setValues(newRows);
  } finally {
    lock.releaseLock();
  }

  upsertCustomer_({ openChatName: data.openChatName, lineName: data.lineName });
  return { orderId: orderId, total: total };
}

/**
 * 管理者ページ用: 注文のステータスを更新する(同一注文IDの全行)。
 * 在庫管理: 「承認待ち/キャンセル → 注文確定など」への変更時に在庫を引き落とし
 * (不足していれば例外を投げて変更しない)、逆方向への変更時は在庫を戻す。
 * 承認ボタンは status='注文確定' でこの関数を呼ぶ。
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
    var rowIdxs = [];
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(orderId)) rowIdxs.push(i + 2);
    }
    if (rowIdxs.length === 0) return { updated: 0 };

    var oldStatus = String(sheet.getRange(rowIdxs[0], COL_STATUS).getValue()) || STATUSES[0];
    var wasTaken = STOCK_TAKEN_STATUSES.indexOf(oldStatus) !== -1;
    var willTake = STOCK_TAKEN_STATUSES.indexOf(status) !== -1;

    if (wasTaken !== willTake) {
      var order = findOrder_(orderId);
      if (order) {
        // 引落時に在庫不足ならここで例外 → ステータスは変更されない
        adjustStockForOrder_(order, willTake ? -1 : +1);
      }
    }

    rowIdxs.forEach(function (row) {
      sheet.getRange(row, COL_STATUS).setValue(status);
    });
    return { updated: rowIdxs.length, oldStatus: oldStatus };
  } finally {
    lock.releaseLock();
  }
}

// ===== 内部ヘルパー =====

/** 注文一覧シートを注文ID単位のオブジェクト配列(シート順)として読み込む。 */
function readOrders_() {
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
          items: [],
          weightG: 0, shipMethod: '', shipFee: '',
          recvName: '', recvAddress: '', recvPhone: '',
          trackingNo: '', shippedDate: ''
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
      // 発送関連は先頭行にのみ入っているため、値のある行から拾う
      if (r[12] !== '' && r[12] !== null && !orders[id].weightG) orders[id].weightG = Number(r[12]);
      if (r[13] && !orders[id].shipMethod) orders[id].shipMethod = String(r[13]);
      if (r[14] !== '' && r[14] !== null && orders[id].shipFee === '') orders[id].shipFee = Number(r[14]);
      if (r[15] && !orders[id].recvName) orders[id].recvName = String(r[15]);
      if (r[16] && !orders[id].recvAddress) orders[id].recvAddress = String(r[16]);
      if (r[17] && !orders[id].recvPhone) orders[id].recvPhone = String(r[17]);
      if (r[18] && !orders[id].trackingNo) orders[id].trackingNo = String(r[18]);
      if (r[19] && !orders[id].shippedDate) {
        orders[id].shippedDate = r[19] instanceof Date
          ? Utilities.formatDate(r[19], TIMEZONE, 'yyyy/MM/dd') : String(r[19]);
      }
    });
  }
  return orderIdsInOrder.map(function (id) { return orders[id]; });
}

/** 注文番号で1件の注文を探す。修正画面用に備考から修正マークを除去して返す。 */
function findOrder_(orderId) {
  var target = String(orderId || '').trim();
  if (!target) return null;
  var list = readOrders_();
  for (var i = 0; i < list.length; i++) {
    if (list[i].orderId === target) {
      list[i].note = list[i].note.replace(/^【[^】]*修正】\s*/, '');
      return list[i];
    }
  }
  return null;
}

function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('シート「' + name + '」が見つかりません。メニューの「初期セットアップ」を実行してください。');
  }
  return sheet;
}

/** 商品マスターから有効な商品を表示順で返す。stockが '' の商品は在庫管理しない。 */
function getActiveProducts_() {
  var sheet = getSheet_(SHEET_PRODUCTS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var products = [];
  values.forEach(function (r) {
    var id = String(r[0]).trim();
    var name = String(r[1]).trim();
    if (!id || !name) return;
    var active = String(r[4]).trim();
    if (active === '無効' || active === 'FALSE' || active === 'false' || r[4] === false) return;
    var price = (r[2] === '' || r[2] === null || isNaN(Number(r[2]))) ? '' : Number(r[2]);
    var stock = (r[5] === '' || r[5] === null || isNaN(Number(r[5]))) ? '' : Math.floor(Number(r[5]));
    products.push({
      id: id,
      name: name,
      price: price,
      stock: stock,
      sortOrder: (r[3] === '' || isNaN(Number(r[3]))) ? 9999 : Number(r[3])
    });
  });
  products.sort(function (a, b) { return a.sortOrder - b.sortOrder; });
  return products;
}

// ===== 顧客マスター =====

/** 顧客マスターシートを取得する(なければヘッダー付きで自動作成)。 */
function ensureCustomerSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CUSTOMERS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CUSTOMERS);
    sheet.getRange(1, 1, 1, 9).setValues([[
      '顧客No', 'オープンチャット名', 'LINEアカウント名',
      'お届け先氏名', 'お届け先住所', 'お届け先電話',
      '初回注文日', '最終注文日', 'メモ'
    ]]).setFontWeight('bold').setBackground('#ead1dc');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** 顧客マスターの全行を返す。 */
function readCustomers_() {
  var sheet = ensureCustomerSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  var list = [];
  values.forEach(function (r, i) {
    var no = String(r[0]).trim();
    var name = String(r[1]).trim();
    if (!no || !name) return;
    list.push({
      row: i + 2,
      custNo: no,
      openChatName: name,
      lineName: String(r[2] || ''),
      recvName: String(r[3] || ''),
      recvAddress: String(r[4] || ''),
      recvPhone: String(r[5] || ''),
      firstOrderAt: r[6] instanceof Date ? Utilities.formatDate(r[6], TIMEZONE, 'yyyy/MM/dd') : String(r[6] || ''),
      lastOrderAt: r[7] instanceof Date ? Utilities.formatDate(r[7], TIMEZONE, 'yyyy/MM/dd') : String(r[7] || '')
    });
  });
  return list;
}

/**
 * 注文情報から顧客マスターを自動登録・更新する(オープンチャット名で照合)。
 * 新規なら顧客Noを自動採番。既存なら空欄でない項目だけ上書きし最終注文日を更新。
 * 顧客マスターの不具合で注文自体を失敗させないよう、エラーは握りつぶしてログに残す。
 */
function upsertCustomer_(info) {
  try {
    var name = String(info.openChatName || '').trim();
    if (!name) return;
    var lock = LockService.getScriptLock();
    lock.waitLock(10 * 1000);
    try {
      var sheet = ensureCustomerSheet_();
      var today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy/MM/dd');
      var customers = readCustomers_();
      var hit = null;
      customers.forEach(function (c) {
        if (c.openChatName === name) hit = c;
      });

      if (hit) {
        var updates = { 3: info.lineName, 4: info.recvName, 5: info.recvAddress, 6: info.recvPhone };
        for (var col in updates) {
          var v = String(updates[col] || '').trim();
          if (v) sheet.getRange(hit.row, Number(col)).setValue(v);
        }
        sheet.getRange(hit.row, 8).setValue(today); // 最終注文日
      } else {
        // 既存の最大番号+1で採番(C001形式)
        var maxNo = 0;
        customers.forEach(function (c) {
          var m = c.custNo.toUpperCase().match(/^C(\d+)$/);
          if (m && Number(m[1]) > maxNo) maxNo = Number(m[1]);
        });
        var next = String(maxNo + 1);
        var custNo = 'C' + (next.length >= 3 ? next : ('00' + next).slice(-3));
        sheet.appendRow([
          custNo, name,
          String(info.lineName || '').trim(),
          String(info.recvName || '').trim(),
          String(info.recvAddress || '').trim(),
          String(info.recvPhone || '').trim(),
          today, today, ''
        ]);
      }
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    console.error('顧客マスター更新エラー: ' + err);
  }
}

/**
 * 管理者専用: 顧客Noで顧客情報を読み出す(代理入力フォーム用)。
 * 暗号キー(管理者キー)がなければ一切返さない — 個人情報保護のため。
 */
function getCustomer(key, custNo) {
  requireAdmin_(key);
  var no = String(custNo || '').trim().toUpperCase();
  if (/^\d+$/.test(no)) {
    no = 'C' + (no.length >= 3 ? no : ('00' + no).slice(-3)); // 「12」→「C012」も許容
  }
  if (!no) throw new Error('顧客Noを入力してください。');

  var customers = readCustomers_();
  for (var i = 0; i < customers.length; i++) {
    if (customers[i].custNo.toUpperCase() === no) {
      var c = customers[i];
      return {
        custNo: c.custNo,
        openChatName: c.openChatName,
        lineName: c.lineName,
        recvName: c.recvName,
        recvAddress: c.recvAddress,
        recvPhone: c.recvPhone
      };
    }
  }
  throw new Error('顧客No「' + no + '」が見つかりません。管理者ページの「顧客」タブで番号をご確認ください。');
}

/** 商品マスターの全行を行番号付きで返す(在庫更新用)。 */
function readProductsRaw_() {
  var sheet = getSheet_(SHEET_PRODUCTS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var list = [];
  values.forEach(function (r, i) {
    var name = String(r[1]).trim();
    if (!name) return;
    var stock = (r[5] === '' || r[5] === null || isNaN(Number(r[5]))) ? '' : Math.floor(Number(r[5]));
    list.push({ row: i + 2, name: name, stock: stock });
  });
  return list;
}

/**
 * 注文の明細に応じて在庫を増減する(sign: -1=引落 / +1=戻し)。
 * 在庫管理対象(商品マスターに同名があり在庫数が数値)の商品のみが対象。
 * 引落時に在庫が足りない場合は例外を投げ、一切書き込まない。
 */
function adjustStockForOrder_(order, sign) {
  var raw = readProductsRaw_();
  var byName = {};
  raw.forEach(function (p) { byName[p.name] = p; });

  // 同一商品が複数行に分かれている場合に備えて数量を合算
  var need = {};
  order.items.forEach(function (it) {
    var p = byName[it.name];
    if (!p || p.stock === '') return; // 自由入力・在庫管理なしは対象外
    need[it.name] = (need[it.name] || 0) + it.qty;
  });

  var shortages = [];
  Object.keys(need).forEach(function (name) {
    if (sign < 0 && byName[name].stock < need[name]) {
      shortages.push(name + '(残り' + byName[name].stock + '点/注文' + need[name] + '点)');
    }
  });
  if (shortages.length > 0) {
    throw new Error('在庫不足のため承認できません: ' + shortages.join('、') +
      '。注文を修正するか、商品マスターの在庫数を確認してください。');
  }

  var sheet = getSheet_(SHEET_PRODUCTS);
  Object.keys(need).forEach(function (name) {
    var p = byName[name];
    sheet.getRange(p.row, 6).setValue(p.stock + sign * need[name]);
  });
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

/**
 * 送料マスターを読み込む。
 * @return {Array} [{method: 'EMS', tiers: [{maxG: 500, fee: 450}, ...]}, ...]
 */
function getShippingRates_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SHIPPING);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var byMethod = {};
  var methodOrder = [];
  values.forEach(function (r) {
    var method = String(r[0]).trim();
    var maxG = Number(r[1]);
    var fee = Number(r[2]);
    if (!method || !(maxG > 0) || isNaN(fee)) return;
    if (!byMethod[method]) {
      byMethod[method] = [];
      methodOrder.push(method);
    }
    byMethod[method].push({ maxG: maxG, fee: fee });
  });
  return methodOrder.map(function (m) {
    byMethod[m].sort(function (a, b) { return a.maxG - b.maxG; });
    return { method: m, tiers: byMethod[m] };
  });
}

/** 配送方法と重量(g)から送料を計算する。該当帯がなければ null。 */
function calcShippingFee_(method, weightG) {
  var rates = getShippingRates_();
  for (var i = 0; i < rates.length; i++) {
    if (rates[i].method !== method) continue;
    var tiers = rates[i].tiers;
    for (var j = 0; j < tiers.length; j++) {
      if (weightG <= tiers[j].maxG) return tiers[j].fee;
    }
    return null; // 最大重量帯を超過
  }
  return null; // 未知の配送方法
}

/**
 * 注文の先頭行に指定フィールドを書き込み、statusが指定されていれば全行のステータスを更新する。
 * @param {Object} fields {列番号(1始まり): 値}
 */
function setOrderFields_(orderId, fields, status) {
  if (status && STATUSES.indexOf(status) === -1) throw new Error('不正なステータスです。');
  var lock = LockService.getScriptLock();
  lock.waitLock(20 * 1000);
  try {
    var sheet = getSheet_(SHEET_ORDERS);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('注文番号「' + orderId + '」は見つかりません。');
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var firstRow = 0;
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(orderId)) {
        if (!firstRow) firstRow = i + 2;
        if (status) sheet.getRange(i + 2, COL_STATUS).setValue(status);
      }
    }
    if (!firstRow) throw new Error('注文番号「' + orderId + '」は見つかりません。');
    for (var col in fields) {
      sheet.getRange(firstRow, Number(col)).setValue(fields[col]);
    }
  } finally {
    lock.releaseLock();
  }
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

  var p = ss.getSheetByName(SHEET_PRODUCTS);
  if (!p) {
    p = ss.insertSheet(SHEET_PRODUCTS);
    p.getRange(2, 1, 2, 6).setValues([
      ['P001', 'サンプル商品A', 120, 1, '有効', 10],
      ['P002', 'サンプル商品B', 250, 2, '有効', '']
    ]);
    p.setFrozenRows(1);
  }
  // ヘッダーは常に書き直す(旧バージョンからの在庫数列の追加にも対応)
  p.getRange(1, 1, 1, 6).setValues([['商品ID', '商品名', '価格(元)', '表示順', '有効', '在庫数']])
    .setFontWeight('bold').setBackground('#d9ead3');

  var o = ss.getSheetByName(SHEET_ORDERS);
  if (!o) {
    o = ss.insertSheet(SHEET_ORDERS);
    o.setFrozenRows(1);
  }
  // ヘッダーは常に書き直す(旧バージョンからの列追加にも対応)
  o.getRange(1, 1, 1, ORDER_HEADERS.length).setValues([ORDER_HEADERS])
    .setFontWeight('bold').setBackground('#cfe2f3');
  // ステータス列に入力規則(プルダウン)を設定(ステータス追加にも対応して毎回更新)
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build();
  o.getRange(2, COL_STATUS, o.getMaxRows() - 1, 1).setDataValidation(rule);

  ensureCustomerSheet_();

  if (!ss.getSheetByName(SHEET_SHIPPING)) {
    var sh = ss.insertSheet(SHEET_SHIPPING);
    sh.getRange(1, 1, 1, 3).setValues([['配送方法', '重量上限(g)', '料金(元)']])
      .setFontWeight('bold').setBackground('#d9d2e9');
    // ※サンプル料金です。必ず日本郵便の台湾宛料金(円)を台湾元に換算して更新してください。
    var rows = [];
    for (var w = 500, fee = 450; w <= 10000; w += 500, fee += 80) {
      rows.push(['EMS', w, fee]);
    }
    for (var w2 = 1000, fee2 = 450; w2 <= 10000; w2 += 1000, fee2 += 120) {
      rows.push(['国際小包(航空)', w2, fee2]);
    }
    sh.getRange(2, 1, rows.length, 3).setValues(rows);
    sh.setFrozenRows(1);
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
    '2.「商品マスター」シートに商品と在庫数を登録してください。\n' +
    '  (在庫数が空欄の商品は在庫管理なし=無制限になります)\n' +
    '3.「送料マスター」の料金はサンプルです。必ず日本郵便の\n' +
    '   台湾宛料金表(EMS/国際小包)を台湾元に換算して更新してください。\n' +
    '4. デプロイしてURLを取得してください(SETUP.md参照)。'
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

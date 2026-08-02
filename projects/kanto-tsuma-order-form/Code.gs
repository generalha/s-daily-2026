/**
 * 關東太太 注文・発送・利益管理システム セットアップスクリプト
 * ------------------------------------------------------------
 * このスクリプトを1回実行するだけで、
 *   ・注文者（お客様）が直接入力するGoogleフォーム
 *   ・回答が自動で貯まるスプレッドシート（回答／設定／ダッシュボードの3シート）
 * をまとめて自動生成します。
 *
 * 使い方は同じフォルダの README.md を参照してください。
 *
 * 【運用の考え方（超重要）】
 * ・フォームは「注文者（お客様）」が入力するもの。仕入れ値や支払い状況など
 *   社内管理の情報はフォームに含めない（お客様に見せない・入力させない）。
 * ・仕入れ価格・国内送料・支払い状況・発送状況・発送方法・重量・追跡番号は、
 *   注文が入った後に妻が『回答』シートのI〜O列に直接入力する（プルダウン付き）。
 * ・注文ID、円換算売上、粗利益、利益率、EMS用コピペ、案内文の下書きは
 *   すべて数式（ARRAYFORMULA）で自動計算されるので、
 *   トリガーの設定やスクリプトの実行は最初の1回だけでOK。
 * ・住所・電話番号など機微情報の項目はこのフォームには含めていません。
 * ・EMSラベルへの自動入力は行いません。「EMS用コピペ欄」をコピーして
 *   手作業でEMSの発送手続き画面に貼り付ける想定です。
 */

// ============================================================
// 設定値（必要ならここだけ書き換えてOK）
// ============================================================

var FORM_TITLE = '關東太太 注文・発送管理フォーム';
var FORM_DESCRIPTION =
  'ご注文内容の確認用フォームです。ご入力いただいた内容を確認のうえ、' +
  'お支払い・発送についてこちらからあらためてご案内いたします。';

var SHEET_NAME_RESPONSES = '回答';       // フォームの回答が入るシート
var SHEET_NAME_SETTINGS = '設定';        // 為替レートなどの設定シート
var SHEET_NAME_DASHBOARD = 'ダッシュボード'; // 集計・分析シート

var DEFAULT_EXCHANGE_RATE = 0.25; // 初期為替レート（意味は README・設定シート参照）

// ============================================================
// メイン関数：これを実行すればフォーム一式が作られます
// ============================================================

/**
 * Apps Scriptエディタでこの関数を選んで「実行」を押してください。
 * 初回のみ実行し、2回目以降は既存のフォーム／シートのURLをログに出すだけにします
 * （誤って2回実行しても、フォームが重複して作られないようにするためです）。
 */
function setupKantoTsumaSystem() {
  var props = PropertiesService.getScriptProperties();
  var existingFormId = props.getProperty('KANTO_TSUMA_FORM_ID');
  var existingSheetId = props.getProperty('KANTO_TSUMA_SHEET_ID');

  if (existingFormId && existingSheetId) {
    var existingForm = FormApp.openById(existingFormId);
    var existingSheet = SpreadsheetApp.openById(existingSheetId);
    Logger.log('すでにセットアップ済みです。以下のURLを使ってください。');
    Logger.log('フォーム編集URL: ' + existingForm.getEditUrl());
    Logger.log('フォーム回答URL(注文者に共有する画面): ' + existingForm.getPublishedUrl());
    Logger.log('スプレッドシートURL: ' + existingSheet.getUrl());
    return;
  }

  // 1. フォームとスプレッドシートを作成してひも付け
  var form = createOrderForm_();
  var ss = SpreadsheetApp.create(FORM_TITLE + ' 回答');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // 2. フォームの回答シートを見つけて「回答」という名前にそろえる
  var responseSheet = findFormResponseSheet_(ss);
  responseSheet.setName(SHEET_NAME_RESPONSES);

  // 3. 回答シートに「妻が入力する管理列」と「自動計算列」を追加
  setupStaffManagedColumns_(responseSheet);
  setupComputedColumns_(responseSheet);

  // 4. 設定シートを作成
  var settingsSheet = ss.insertSheet(SHEET_NAME_SETTINGS);
  setupSettingsSheet_(settingsSheet);

  // 5. ダッシュボードシートを作成
  var dashboardSheet = ss.insertSheet(SHEET_NAME_DASHBOARD);
  setupDashboardSheet_(dashboardSheet);

  // 6. 不要な初期シート（Sheet1／シート1など）を削除
  removeDefaultBlankSheets_(ss, [SHEET_NAME_RESPONSES, SHEET_NAME_SETTINGS, SHEET_NAME_DASHBOARD]);

  // 7. シートの並び順を「回答→設定→ダッシュボード」にそろえる
  responseSheet.activate();
  ss.moveActiveSheet(1);
  settingsSheet.activate();
  ss.moveActiveSheet(2);
  dashboardSheet.activate();
  ss.moveActiveSheet(3);
  responseSheet.activate();

  // 8. 次回以降の誤操作防止のため、作成したIDを保存
  props.setProperty('KANTO_TSUMA_FORM_ID', form.getId());
  props.setProperty('KANTO_TSUMA_SHEET_ID', ss.getId());

  // 9. 完成したURLをログに出力
  Logger.log('===== セットアップ完了 =====');
  Logger.log('フォーム編集URL: ' + form.getEditUrl());
  Logger.log('フォーム回答URL(注文者に共有する画面): ' + form.getPublishedUrl());
  Logger.log('スプレッドシートURL: ' + ss.getUrl());
}

// ============================================================
// 1. フォーム作成（注文者＝お客様が直接入力する項目のみ）
// ============================================================

function createOrderForm_() {
  var form = FormApp.create(FORM_TITLE);
  form.setDescription(FORM_DESCRIPTION);
  form.setCollectEmail(false); // メールアドレス収集はしない（機微情報を増やさない）

  // ---- 注文者名（短文・必須） ----
  form.addTextItem()
    .setTitle('注文者名')
    .setRequired(true);

  // ---- 販売先（プルダウン・必須） ----
  form.addListItem()
    .setTitle('販売先（どちらでご注文いただきましたか）')
    .setChoiceValues(['LINEオープンチャット', '個別LINE', 'Facebook', 'その他'])
    .setRequired(true);

  // ---- 商品カテゴリー（プルダウン・必須） ----
  form.addListItem()
    .setTitle('商品カテゴリー')
    .setChoiceValues(['戰鬥陀螺', '童鞋', '周邊商品', '食品', 'その他'])
    .setRequired(true);

  // ---- 商品名・型番（短文・必須） ----
  form.addTextItem()
    .setTitle('商品名・型番')
    .setRequired(true);

  // ---- 数量（数値・必須） ----
  form.addTextItem()
    .setTitle('数量')
    .setRequired(true)
    .setValidation(numberValidation_());

  // ---- 販売価格（台湾ドル）（数値・必須） ----
  // ご案内済みの金額をそのままご入力いただく想定です。
  form.addTextItem()
    .setTitle('販売価格（台湾ドル）')
    .setRequired(true)
    .setValidation(numberValidation_());

  // ---- メモ（段落・任意） ----
  form.addParagraphTextItem()
    .setTitle('メモ（ご要望などあればご記入ください）')
    .setRequired(false);

  return form;
}

/** 数値入力チェック（半角数字以外を入力するとエラーメッセージが出る） */
function numberValidation_() {
  return FormApp.createTextValidation()
    .setHelpText('半角数字で入力してください。')
    .requireNumber()
    .build();
}

// ============================================================
// 2. 回答シート：妻が入力する管理列（I〜O）のセットアップ
// ============================================================

/**
 * フォームの質問は A〜H列に自動で入る（A列はタイムスタンプ）。
 *   A タイムスタンプ
 *   B 注文者名
 *   C 販売先
 *   D 商品カテゴリー
 *   E 商品名・型番
 *   F 数量
 *   G 販売価格（台湾ドル）
 *   H メモ
 *
 * この関数はその右側 I〜O列に、仕入れ値・送料・支払い状況・発送状況・
 * 発送方法・重量・追跡番号の「妻が注文確認後に入力する管理欄」を追加する。
 * 支払い状況／発送状況／発送方法はプルダウンにしてあるので、
 * クリックして選ぶだけで入力できる（入力ミス防止）。
 * 見分けやすいように背景色を薄い黄色にしている。
 */
function setupStaffManagedColumns_(sheet) {
  var headers = [
    '仕入れ価格（日本円）', '国内送料（日本円）', '支払い状況',
    '発送状況', '発送方法', '重量（グラム）', '追跡番号'
  ];
  sheet.getRange(1, 9, 1, headers.length).setValues([headers]); // I1〜O1

  var lastRow = sheet.getMaxRows();

  // 支払い状況（K列）のプルダウン
  var paymentRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未請求', '請求済み', '一部入金', '入金済み'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 11, lastRow - 1, 1).setDataValidation(paymentRule); // K2:K

  // 発送状況（L列）のプルダウン
  var shippingRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未発送', '同梱待ち', '梱包済み', '発送済み'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 12, lastRow - 1, 1).setDataValidation(shippingRule); // L2:L

  // 発送方法（M列）のプルダウン
  var methodRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未定', 'EMS', '国際小包航空便', 'その他'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 13, lastRow - 1, 1).setDataValidation(methodRule); // M2:M

  // 「妻が入力する欄」だとひと目で分かるよう薄い黄色を付ける
  sheet.getRange(1, 9, lastRow, 7).setBackground('#fff2cc');
}

// ============================================================
// 3. 回答シート：自動計算列（P〜V）のセットアップ
// ============================================================

/**
 * P〜V列に、注文ID・円換算売上・粗利益・利益率・EMS用コピペ・
 * 案内文の下書きをARRAYFORMULA（数式）で自動追加する。
 * ARRAYFORMULAを2行目に1つ入れておけば、フォームで新しい回答が
 * 追加されるたびに自動で計算されるので、その都度スクリプトを
 * 動かす必要はない。
 */
function setupComputedColumns_(sheet) {
  var headers = [
    '注文ID', '為替レート（使用値）', '売上（円換算）', '粗利益（円）', '利益率',
    'EMS用コピペ', '案内文の下書き'
  ];
  sheet.getRange(1, 16, 1, headers.length).setValues([headers]); // P1〜V1

  // 注文ID：KT-0001 のような連番
  sheet.getRange('P2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","","KT-"&TEXT(ROW(A2:A)-1,"0000")))'
  );

  // 為替レート（使用値）：設定シートの値をそのまま反映
  sheet.getRange('Q2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",' + SHEET_NAME_SETTINGS + '!$B$2))'
  );

  // 売上（円換算）＝ 販売価格（台湾ドル） ÷ 為替レート
  sheet.getRange('R2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",IFERROR(G2:G/Q2:Q,"")))'
  );

  // 粗利益（円）＝ 売上（円換算） − 仕入れ価格(I) − 国内送料(J)（未入力は0扱い）
  sheet.getRange('S2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",IFERROR(R2:R-N(I2:I)-N(J2:J),"")))'
  );

  // 利益率 ＝ 粗利益 ÷ 売上（円換算）
  sheet.getRange('T2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",IFERROR(S2:S/R2:R,"")))'
  );

  // EMS用コピペ欄：注文者・商品名・数量・発送方法(M)・重量(N)・販売価格(G)・メモ(H)を1行に整形
  sheet.getRange('U2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",' +
      'B2:B&" ｜ "&E2:E&" x"&F2:F&"個 ｜ 発送方法:"&IF(M2:M="","未定",M2:M)&' +
      '" ｜ 重量:"&IF(N2:N="","未計測",N2:N&"g")&" ｜ 販売価格:NT$"&G2:G&' +
      '" ｜ メモ:"&IF(H2:H="","なし",H2:H)))'
  );

  // 案内文の下書き：お客さん向けにそのままコピペできる文面
  sheet.getRange('V2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",' +
      'B2:B&"様"&CHAR(10)&CHAR(10)&' +
      '"この度はご注文いただきありがとうございます。"&CHAR(10)&' +
      '"内容を確認しましたのでご案内いたします。"&CHAR(10)&CHAR(10)&' +
      '"商品名："&E2:E&CHAR(10)&' +
      '"数量："&F2:F&"個"&CHAR(10)&' +
      '"商品代金：NT$"&G2:G&CHAR(10)&' +
      '"発送方法："&IF(M2:M="","未定",M2:M)&CHAR(10)&' +
      '"送料につきましては、梱包完了後にあらためてご案内いたします。"&CHAR(10)&CHAR(10)&' +
      '"引き続きどうぞよろしくお願いいたします。"))'
  );

  // 見た目を整える
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 22).setFontWeight('bold'); // A〜V見出しを太字
  sheet.setColumnWidth(21, 320); // U列（EMS用コピペ）を広めに
  sheet.setColumnWidth(22, 380); // V列（案内文）を広めに
  sheet.getRange('V2:V').setWrap(true);
  sheet.getRange('U2:U').setWrap(true);
  sheet.getRange('T2:T').setNumberFormat('0.0%');
  sheet.getRange('R2:R').setNumberFormat('¥#,##0');
  sheet.getRange('S2:S').setNumberFormat('¥#,##0');

  // 数式が入っているP〜V列を誤って上書きしないよう「警告」を出す
  // （ブロックはしない＝奥さんが困らないよう、警告のみに留める）
  var protectedRange = sheet.getRange(2, 16, sheet.getMaxRows() - 1, 7); // P2:V最終行
  var protection = protectedRange.protect();
  protection.setWarningOnly(true);
  protection.setDescription('自動計算欄です。数式を消さないよう注意してください。');
}

/** SpreadsheetApp.create直後にフォームからひも付けられた回答シートを探す */
function findFormResponseSheet_(ss) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.indexOf('フォームの回答') === 0 || name.indexOf('Form Responses') === 0) {
      return sheets[i];
    }
  }
  // 万が一見つからない場合は最初のシートを使う
  return sheets[0];
}

/** セットアップ用に作られた空の初期シート（Sheet1／シート1）を削除する */
function removeDefaultBlankSheets_(ss, keepNames) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (keepNames.indexOf(name) === -1) {
      ss.deleteSheet(sheets[i]);
    }
  }
}

// ============================================================
// 4. 設定シート
// ============================================================

function setupSettingsSheet_(sheet) {
  sheet.getRange('A1').setValue('關東太太 設定シート').setFontWeight('bold').setFontSize(14);

  sheet.getRange('A2').setValue('為替レート（1円 ＝ 何台湾ドルか）');
  sheet.getRange('B2').setValue(DEFAULT_EXCHANGE_RATE);
  sheet.getRange('B2').setNumberFormat('0.0000');

  sheet.getRange('A4').setValue('使い方メモ').setFontWeight('bold');
  sheet.getRange('A5').setValue(
    'B2セルの数値を変えると、回答シートの「売上（円換算）」「粗利益」「利益率」が\n' +
    '自動で再計算されます（フォーマットを直す必要はありません）。'
  );
  sheet.getRange('A6').setValue(
    '計算式：売上（円）＝ 販売価格（台湾ドル）÷ B2の値'
  );
  sheet.getRange('A7').setValue(
    '例：1台湾ドル＝約4.6円のとき、B2は 1 ÷ 4.6 ≒ 0.217 になります。' +
    '台湾の銀行やGoogleで「TWD JPY レート」を検索し、月1回程度の頻度で更新してください。'
  );
  sheet.getRange('A8').setValue(
    '※ 送金・両替の実勢レートより少しだけ厳しめ（円換算が少なめに出る）の数値にしておくと、' +
    '利益を過大に見積もらずに済みます。'
  );

  sheet.getRange('A5:A8').setWrap(true);
  sheet.setColumnWidth(1, 620);
  sheet.setColumnWidth(2, 120);
}

// ============================================================
// 5. ダッシュボードシート
// ============================================================

function setupDashboardSheet_(sheet) {
  var R = SHEET_NAME_RESPONSES;

  sheet.getRange('A1').setValue('關東太太 ダッシュボード').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A2').setValue('※ すべて数式なので自動更新されます。手動で数字を入れないでください。')
    .setFontColor('#999999');

  // ---- サマリー ----
  sheet.getRange('A4').setValue('総売上（円換算）');
  sheet.getRange('B4').setFormula('=SUM(' + R + '!R2:R)');
  sheet.getRange('B4').setNumberFormat('¥#,##0');

  sheet.getRange('A5').setValue('総粗利益（円）');
  sheet.getRange('B5').setFormula('=SUM(' + R + '!S2:S)');
  sheet.getRange('B5').setNumberFormat('¥#,##0');

  // 未入金＝支払い状況が「入金済み」以外（未記入の新規注文も含む）
  sheet.getRange('A6').setValue('未入金件数（入金済み以外）');
  sheet.getRange('B6').setFormula(
    '=COUNTIFS(' + R + '!B2:B,"<>",' + R + '!K2:K,"<>入金済み")'
  );

  // 未発送＝発送状況が「発送済み」以外（未記入の新規注文も含む）
  sheet.getRange('A7').setValue('未発送件数（発送済み以外）');
  sheet.getRange('B7').setFormula(
    '=COUNTIFS(' + R + '!B2:B,"<>",' + R + '!L2:L,"<>発送済み")'
  );

  sheet.getRange('A4:A7').setFontWeight('bold');

  // ---- 商品別売上ランキング ----
  sheet.getRange('A9').setValue('商品別 売上ランキング（円換算）').setFontWeight('bold');
  sheet.getRange('A10').setFormula(
    '=IFERROR(QUERY(' + R + '!A2:V,' +
    '"select E, sum(R) where E is not null group by E order by sum(R) desc ' +
    'label E \'商品名・型番\', sum(R) \'売上合計（円）\'"),"データがまだありません")'
  );

  // ---- 商品別粗利益ランキング ----
  sheet.getRange('D9').setValue('商品別 粗利益ランキング（円）').setFontWeight('bold');
  sheet.getRange('D10').setFormula(
    '=IFERROR(QUERY(' + R + '!A2:V,' +
    '"select E, sum(S) where E is not null group by E order by sum(S) desc ' +
    'label E \'商品名・型番\', sum(S) \'粗利益合計（円）\'"),"データがまだありません")'
  );

  sheet.setColumnWidth(1, 260);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(4, 260);
  sheet.setColumnWidth(5, 150);
}

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, VerticalAlign, ShadingType,
  PageBreak, HeadingLevel,
} = require("docx");
const fs = require("fs");

const FONT = { ascii: "MS Mincho", hAnsi: "MS Mincho", eastAsia: "ＭＳ 明朝" };
const GOTHIC = { ascii: "MS Gothic", hAnsi: "MS Gothic", eastAsia: "ＭＳ ゴシック" };

const CONTENT_W = 9638; // A4 11906 - margins 1134*2

function t(text, opts = {}) {
  return new TextRun({ text, font: opts.gothic ? GOTHIC : FONT, size: opts.size || 20, bold: opts.bold || false });
}
function p(text, opts = {}) {
  return new Paragraph({
    alignment: opts.align,
    spacing: { before: opts.before ?? 40, after: opts.after ?? 40, line: opts.line },
    children: Array.isArray(text) ? text : [t(text, opts)],
  });
}
function sectionHeader(text, opts = {}) {
  return new Paragraph({
    pageBreakBefore: opts.pageBreakBefore || false,
    spacing: { before: 160, after: 60 },
    children: [t(text, { gothic: true, bold: true, size: 21 })],
  });
}
const thinBorders = {
  top: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
  left: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
  right: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
};
function cell(children, width, opts = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    shading: opts.shade ? { type: ShadingType.CLEAR, fill: "EEEEEE" } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    columnSpan: opts.span,
    children: (Array.isArray(children) ? children : [children]).map((c) =>
      typeof c === "string" ? p(c, { before: 0, after: 0, size: opts.size || 20, align: opts.align }) : c
    ),
  });
}
function table(columnWidths, rows) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths,
    borders: thinBorders,
    rows,
  });
}
function row(cells) { return new TableRow({ children: cells }); }

const children = [];

// ===== Page 1 =====
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 120 },
  children: [t("保育料（利用者負担額）算定に係る生計状況申立書", { gothic: true, bold: true, size: 30 })],
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 160 },
  children: [t("（同居の祖父母等と生計を別にしていることの申立て）", { gothic: true, size: 22 })],
}));

children.push(p("　　　　　　　　　市（町・村）長　様", { size: 22 }));
children.push(p("　", { size: 8 }));
children.push(new Paragraph({
  alignment: AlignmentType.RIGHT,
  spacing: { after: 60 },
  children: [t("申立年月日　　令和　　　年　　　月　　　日", { size: 20 })],
}));

children.push(p([
  t("　私は、下記の児童の保育料（利用者負担額）の算定にあたり、同居している祖父母等と", { size: 20 }),
  t("生計を別にしている", { size: 20, bold: true }),
  t("ため、次のとおり申し立てます。", { size: 20 }),
], { before: 60, after: 120 }));

// 申立人・児童
children.push(sectionHeader("１　申立人（保護者）及び対象児童"));
children.push(table(
  [1600, 3219, 1600, 3219],
  [
    row([
      cell("申立人氏名", 1600, { shade: true }),
      cell("", 3219),
      cell("児童との続柄", 1600, { shade: true }),
      cell("", 3219),
    ]),
    row([
      cell("住　所", 1600, { shade: true }),
      cell("", 3219, { span: 3 }),
    ]),
    row([
      cell("電話番号", 1600, { shade: true }),
      cell("", 3219),
      cell("児童氏名", 1600, { shade: true }),
      cell("", 3219),
    ]),
    row([
      cell("児童生年月日", 1600, { shade: true }),
      cell("平成・令和　　年　　月　　日", 3219, { size: 18 }),
      cell("利用（希望）施設", 1600, { shade: true, size: 18 }),
      cell("", 3219),
    ]),
  ]
));

// 同居親族
children.push(sectionHeader("２　同居している祖父母等（父母以外の同居親族）の状況"));
const memberHeader = row([
  cell("氏　名", 2500, { shade: true, align: AlignmentType.CENTER }),
  cell("続柄※", 1500, { shade: true, align: AlignmentType.CENTER }),
  cell("生年月日", 2800, { shade: true, align: AlignmentType.CENTER }),
  cell("職業・就労状況", 2838, { shade: true, align: AlignmentType.CENTER }),
]);
const memberRows = [];
for (let i = 0; i < 3; i++) {
  memberRows.push(row([
    cell("", 2500),
    cell("", 1500),
    cell("昭和・平成　　年　　月　　日", 2800, { size: 18 }),
    cell("", 2838),
  ]));
}
children.push(table([2500, 1500, 2800, 2838], [memberHeader, ...memberRows]));
children.push(p("※　続柄は児童から見た関係（祖父、祖母、曾祖母、おじ　など）を記入してください。", { size: 16, before: 20, after: 40 }));

// 世帯の状況
children.push(sectionHeader("３　住民票上の世帯の状況（該当する□にレ点を記入してください。以下同じ。）"));
children.push(p("□　祖父母等と同一世帯である　　　□　世帯分離している（分離年月日：　　　年　　月　　日）", { before: 20, after: 20 }));
children.push(p("※　住民票上の世帯分離のみをもって別生計とは判断されません。以下の項目により実態を確認します。", { size: 18, after: 60 }));

// 住居の状況
children.push(sectionHeader("４　住居の状況"));
children.push(table(
  [2800, 6838],
  [
    row([
      cell("住宅の種別", 2800, { shade: true }),
      cell("□ 持家（名義人：　　　　　　　）　□ 借家・アパート（契約者：　　　　　　　）　□ その他（　　　　　　　）", 6838, { size: 18 }),
    ]),
    row([
      cell("住宅の構造", 2800, { shade: true }),
      cell("□ 完全分離型二世帯住宅　□ 一部共用の二世帯住宅　□ 一般住宅（共用）", 6838, { size: 18 }),
    ]),
    row([
      cell("玄関・台所・浴室", 2800, { shade: true }),
      cell("玄関（□別・□共用）　台所（□別・□共用）　浴室（□別・□共用）", 6838, { size: 18 }),
    ]),
    row([
      cell("光熱水道の契約・メーター", 2800, { shade: true }),
      cell("電気（□別・□共用）　ガス（□別・□共用）　水道（□別・□共用）", 6838, { size: 18 }),
    ]),
  ]
));

// 家計の状況
children.push(sectionHeader("５　家計の状況（それぞれ主に負担している方を記入してください）"));
children.push(table(
  [2400, 2600, 2038, 2600],
  [
    row([
      cell("費　目", 2400, { shade: true, align: AlignmentType.CENTER }),
      cell("主に負担している者（氏名）", 2600, { shade: true, align: AlignmentType.CENTER, size: 18 }),
      cell("月額（おおよそ）", 2038, { shade: true, align: AlignmentType.CENTER, size: 18 }),
      cell("備考（按分の方法など）", 2600, { shade: true, align: AlignmentType.CENTER, size: 18 }),
    ]),
    row([cell("家賃・住宅ローン", 2400), cell("", 2600), cell("円", 2038, { align: AlignmentType.RIGHT }), cell("", 2600)]),
    row([cell("電気・ガス・水道", 2400), cell("", 2600), cell("円", 2038, { align: AlignmentType.RIGHT }), cell("", 2600)]),
    row([cell("食　費", 2400), cell("", 2600), cell("円", 2038, { align: AlignmentType.RIGHT }), cell("", 2600)]),
    row([cell("通信費（電話等）", 2400), cell("", 2600), cell("円", 2038, { align: AlignmentType.RIGHT }), cell("", 2600)]),
    row([cell("その他（　　　　）", 2400), cell("", 2600), cell("円", 2038, { align: AlignmentType.RIGHT }), cell("", 2600)]),
  ]
));
children.push(p("祖父母等との間の生活費等の金銭のやり取り：□ ない　□ ある（内容：　　　　　　　　　　　　）", { size: 18, before: 60 }));

// ===== Page 2 =====
// 扶養の状況
children.push(sectionHeader("６　扶養等の状況", { pageBreakBefore: true }));
children.push(table(
  [6038, 3600],
  [
    row([
      cell("父母又は児童が、同居の祖父母等の税法上の扶養親族となっている", 6038, { size: 18 }),
      cell("□ なっていない　□ なっている", 3600, { size: 18 }),
    ]),
    row([
      cell("父母又は児童が、同居の祖父母等の健康保険の被扶養者となっている", 6038, { size: 18 }),
      cell("□ なっていない　□ なっている", 3600, { size: 18 }),
    ]),
    row([
      cell("祖父母等が、父母又は児童を対象とする手当・給付等を受給している", 6038, { size: 18 }),
      cell("□ 受給していない　□ 受給している", 3600, { size: 18 }),
    ]),
  ]
));

// 父母の就労・収入
children.push(sectionHeader("７　父母の就労・収入の状況"));
children.push(table(
  [1200, 3400, 2600, 2438],
  [
    row([
      cell("", 1200, { shade: true }),
      cell("勤務先・就労形態", 3400, { shade: true, align: AlignmentType.CENTER, size: 18 }),
      cell("前年の年収（税込）", 2600, { shade: true, align: AlignmentType.CENTER, size: 18 }),
      cell("市町村民税の申告", 2438, { shade: true, align: AlignmentType.CENTER, size: 18 }),
    ]),
    row([
      cell("父", 1200, { align: AlignmentType.CENTER }),
      cell("", 3400),
      cell("約　　　　　　万円", 2600, { align: AlignmentType.RIGHT, size: 18 }),
      cell("□ 済　□ 未申告", 2438, { size: 18 }),
    ]),
    row([
      cell("母", 1200, { align: AlignmentType.CENTER }),
      cell("", 3400),
      cell("約　　　　　　万円", 2600, { align: AlignmentType.RIGHT, size: 18 }),
      cell("□ 済　□ 未申告", 2438, { size: 18 }),
    ]),
  ]
));

// 理由
children.push(sectionHeader("８　生計を別にしている具体的な状況・理由"));
children.push(table(
  [CONTENT_W],
  [
    row([
      new TableCell({
        width: { size: CONTENT_W, type: WidthType.DXA },
        margins: { top: 60, bottom: 60, left: 100, right: 100 },
        children: [p("", { before: 0, after: 0 }), p("", {}), p("", {}), p("", {}), p("", {}), p("", { after: 0 })],
      }),
    ]),
  ]
));

// 添付書類
children.push(sectionHeader("９　添付書類（生計が別であることが確認できる書類。該当するものにレ点）"));
children.push(p("□　公共料金（電気・ガス・水道等）の領収書又は検針票の写し（申立人（父母）名義のもの）", { size: 18, before: 20, after: 20 }));
children.push(p("□　住宅の賃貸借契約書の写し又は住宅ローン返済予定表等の写し（父母が契約・返済しているもの）", { size: 18, after: 20 }));
children.push(p("□　父母及び児童の健康保険証（資格確認書等）の写し", { size: 18, after: 20 }));
children.push(p("□　父母の収入がわかる書類（源泉徴収票、給与明細、確定申告書の控え等の写し）", { size: 18, after: 20 }));
children.push(p("□　その他（　　　　　　　　　　　　　　　　　　　　　　　　　　　　　）", { size: 18, after: 60 }));

// 誓約
children.push(sectionHeader("１０　誓約・同意"));
children.push(p("　上記の記載内容は事実と相違ありません。実態調査等のため、必要に応じて追加書類の提出を求められること、及び公簿等により記載内容が確認されることに同意します。また、申立ての内容が事実と異なることが判明した場合は、保育料（利用者負担額）が再算定され、遡って差額を徴収されることがあることを承知しています。", { size: 20, before: 40, after: 160 }));
children.push(new Paragraph({
  alignment: AlignmentType.RIGHT,
  spacing: { after: 40 },
  children: [
    t("令和　　　年　　　月　　　日　　申立人署名　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿", { size: 22 }),
  ],
}));

// 職員記入欄
children.push(p("　", { size: 8 }));
children.push(table(
  [1800, 2600, 2600, 2638],
  [
    row([
      cell("※職員記入欄", 1800, { shade: true, align: AlignmentType.CENTER, size: 18 }),
      cell("受付日：　　・　　・　　", 2600, { size: 18 }),
      cell("確認書類：", 2600, { size: 18 }),
      cell("判定：□ 別生計　□ 合算", 2638, { size: 18 }),
    ]),
  ]
));

// 記入上の注意
children.push(p("　", { size: 8 }));
children.push(p("【記入にあたっての注意】", { gothic: true, bold: true, size: 18, before: 60, after: 20 }));
children.push(p("・この申立書は、保育料（利用者負担額）の算定において、同居している祖父母等の市町村民税額を合算しないこと（別生計）の判断を受けるために提出するものです。", { size: 16, after: 20 }));
children.push(p("・住民票上世帯分離をしている場合でも、生活実態が同一生計と認められるときは、祖父母等のうち最多収入の方を家計の主宰者として税額が合算されることがあります。", { size: 16, after: 20 }));
children.push(p("・父母の収入状況等によっては、この申立てによらず同一生計と判断される場合があります。詳細は各市町村の担当課にお問い合わせください。", { size: 16, after: 20 }));

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: FONT, size: 20 } },
    },
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1134, bottom: 850, left: 1134, right: 1134 },
        },
      },
      children,
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(process.argv[2] || "form.docx", buf);
  console.log("written");
});

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, VerticalAlign, ShadingType,
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
    spacing: { before: opts.before ?? 90, after: 35 },
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
    margins: { top: 50, bottom: 50, left: 100, right: 100 },
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
  spacing: { after: 100 },
  children: [t("保育料（利用者負担額）算定に係る生計状況申立書", { gothic: true, bold: true, size: 30 })],
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 140 },
  children: [t("（同居の祖父母等と生計を別にしていることの申立て）", { gothic: true, size: 22 })],
}));

children.push(p("　　　　　　　　　市（町・村）長　様", { size: 22 }));
children.push(new Paragraph({
  alignment: AlignmentType.RIGHT,
  spacing: { after: 50 },
  children: [t("申立年月日　　令和　　　年　　　月　　　日", { size: 20 })],
}));

children.push(p([
  t("　私は、下記の児童の保育料（利用者負担額）の算定にあたり、同居している祖父母等と", { size: 20 }),
  t("生計を別にしている", { size: 20, bold: true }),
  t("ため、次のとおり申し立てます。", { size: 20 }),
], { before: 40, after: 100 }));

// 1 申立人・児童
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

// 2 同居親族
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

// 3 住所・世帯 (STEP 1)
children.push(sectionHeader("３　住所・世帯の状況（該当する□にレ点を記入してください。以下同じ。）"));
children.push(table(
  [2800, 6838],
  [
    row([
      cell("住所（地番）の同一性", 2800, { shade: true, size: 18 }),
      cell("祖父母等とは　□ 同一の住所（地番・同一建物）に居住　□ 別の住所（住居表示・地番が異なる）", 6838, { size: 18 }),
    ]),
    row([
      cell("住民票上の世帯", 2800, { shade: true, size: 18 }),
      cell("□ 同一世帯　□ 世帯分離している（分離年月日：　　　年　　月　　日）", 6838, { size: 18 }),
    ]),
  ]
));
children.push(p("※　同居か否かは住所（地番）の同一性により判定され、住民票上の世帯分離のみをもって別生計とは判断されません。", { size: 16, before: 30, after: 40 }));

// 4 父母の課税・就労・収入 (STEP 2・3, 要件D)
children.push(sectionHeader("４　父母の課税・就労・収入の状況"));
children.push(table(
  [900, 2700, 2700, 1800, 1538],
  [
    row([
      cell("", 900, { shade: true }),
      cell("勤務先・就労形態", 2700, { shade: true, align: AlignmentType.CENTER, size: 18 }),
      cell("市町村民税所得割", 2700, { shade: true, align: AlignmentType.CENTER, size: 18 }),
      cell("税の申告", 1800, { shade: true, align: AlignmentType.CENTER, size: 18 }),
      cell("前年の年収", 1538, { shade: true, align: AlignmentType.CENTER, size: 18 }),
    ]),
    row([
      cell("父", 900, { align: AlignmentType.CENTER }),
      cell("", 2700),
      cell("□ 課税　□ 非課税", 2700, { size: 18 }),
      cell("□ 済　□ 未申告", 1800, { size: 18 }),
      cell("約　　　万円", 1538, { align: AlignmentType.RIGHT, size: 18 }),
    ]),
    row([
      cell("母", 900, { align: AlignmentType.CENTER }),
      cell("", 2700),
      cell("□ 課税　□ 非課税", 2700, { size: 18 }),
      cell("□ 済　□ 未申告", 1800, { size: 18 }),
      cell("約　　　万円", 1538, { align: AlignmentType.RIGHT, size: 18 }),
    ]),
    row([
      cell("父母の収入合計（児童手当等の公的手当・養育費を含む）", 8100, { shade: true, size: 18, span: 4 }),
      cell("約　　万円／年", 1538, { align: AlignmentType.RIGHT, size: 16 }),
    ]),
  ]
));
children.push(p("※　未申告のままの場合は、多くの自治体で最高階層により保育料が決定されます。必ず申告を済ませてください。", { size: 16, before: 30, after: 40 }));

// 5 扶養 (STEP 4, 要件C)
children.push(sectionHeader("５　扶養等の状況【要件Ｃ：扶養関係がないこと】"));
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
children.push(p("※　１つでも「なっている・受給している」に該当する場合は、生計同一(合算)と判定され、本申立ては認められません。", { size: 16, before: 30, after: 40 }));

// ===== Page 2 =====
// 6 住居の独立 (要件A)
children.push(sectionHeader("６　住居の状況【要件Ａ：住居の独立】", { pageBreakBefore: true }));
children.push(table(
  [2800, 6838],
  [
    row([
      cell("住宅の種別", 2800, { shade: true }),
      cell("□ 持家（名義人：　　　　　　　）　□ 借家・アパート（契約者：　　　　　　　）　□ その他（　　　　　　　）", 6838, { size: 18 }),
    ]),
    row([
      cell("住宅の構造", 2800, { shade: true }),
      cell("□ 完全分離型二世帯住宅（玄関・台所・浴室すべて別）　□ 一部共用の二世帯住宅　□ 一般住宅（共用）", 6838, { size: 18 }),
    ]),
    row([
      cell("玄関・台所・浴室", 2800, { shade: true }),
      cell("玄関（□別・□共用）　台所（□別・□共用）　浴室（□別・□共用）", 6838, { size: 18 }),
    ]),
    row([
      cell("光熱水道の契約・メーター", 2800, { shade: true, size: 18 }),
      cell("電気（□別・□共用）　ガス（□別・□共用）　水道（□別・□共用）　※「別」は父母名義の契約", 6838, { size: 18 }),
    ]),
  ]
));
children.push(p("※　要件Ａは「完全分離型二世帯住宅」又は「光熱水道の全メーター・契約が父母名義で別」のいずれかで充足します。", { size: 16, before: 30, after: 40 }));

// 7 家計の独立 (要件B)
children.push(sectionHeader("７　家計の状況【要件Ｂ：家計の独立】（それぞれ主に負担している方を記入）"));
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
  ]
));
children.push(p("祖父母等との間の生活費等の金銭のやり取り：□ ない　□ ある（内容：　　　　　　　　　　　　）", { size: 18, before: 25, after: 25 }));

// 8 理由
children.push(sectionHeader("８　生計を別にしている具体的な状況・理由"));
children.push(table(
  [CONTENT_W],
  [
    row([
      new TableCell({
        width: { size: CONTENT_W, type: WidthType.DXA },
        margins: { top: 50, bottom: 50, left: 100, right: 100 },
        children: [p("", { before: 0, after: 0 }), p("", { after: 0 })],
      }),
    ]),
  ]
));

// 9 添付書類
children.push(sectionHeader("９　添付書類（該当するものにレ点。各要件は資料の提出があるものに限り認められます）"));
children.push(p("□　父母名義の公共料金（電気・ガス・水道等）の領収書又は検針票の写し【要件Ａ・Ｂ】", { size: 18, before: 6, after: 6 }));
children.push(p("□　建物図面・登記事項証明書等（完全分離型二世帯住宅の場合）【要件Ａ】", { size: 18, after: 6 }));
children.push(p("□　父母名義の賃貸借契約書の写し又は住宅ローン返済予定表等の写し【要件Ｂ】", { size: 18, after: 6 }));
children.push(p("□　父母及び児童の健康保険証（資格確認書等）の写し【要件Ｃ】", { size: 18, after: 6 }));
children.push(p("□　父母の収入がわかる書類（源泉徴収票・給与明細・確定申告控え等の写し）【要件Ｄ：父母の独立収入】", { size: 18, after: 6 }));
children.push(p("□　その他（　　　　　　　　　　　　　　　　　　　　　）", { size: 18, after: 8 }));

// 10 誓約
children.push(sectionHeader("１０　誓約・同意"));
children.push(p("　上記の記載内容は事実と相違ありません。実態調査等のため、必要に応じて追加書類の提出を求められること、及び公簿等により記載内容が確認されることに同意します。また、申立ての内容が事実と異なることが判明した場合は、保育料（利用者負担額）が再算定され、遡って差額を徴収されることがあることを承知しています。", { size: 20, before: 20, after: 50 }));
children.push(new Paragraph({
  alignment: AlignmentType.RIGHT,
  spacing: { after: 25 },
  children: [t("令和　　　年　　　月　　　日　　申立人署名　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿", { size: 22 })],
}));

// 職員判定欄 (フロー準拠)
children.push(p("※以下は職員記入欄（判定は上から順に行い、結論が出た時点で終了する）", { gothic: true, bold: true, size: 16, before: 40, after: 15 }));
children.push(table(
  [1900, 5100, 2638],
  [
    row([
      cell("STEP１ 同居", 1900, { shade: true, size: 16 }),
      cell("□ 同一地番に居住（同居）　□ 別地番（別居）", 5100, { size: 16 }),
      cell("別居 → 合算しない", 2638, { size: 16 }),
    ]),
    row([
      cell("STEP２ 父母の課税", 1900, { shade: true, size: 16 }),
      cell("□ 所得割課税あり　□ 非課税（申告済）　□ 未申告", 5100, { size: 16 }),
      cell("課税あり → 合算しない／未申告 → 最高階層", 2638, { size: 16 }),
    ]),
    row([
      cell("STEP３ 父母の収入", 1900, { shade: true, size: 16 }),
      cell("基準Ｐ１（　　　　　　円）　□ Ｐ１以上　□ Ｐ１未満", 5100, { size: 16 }),
      cell("Ｐ１以上 → 合算しない", 2638, { size: 16 }),
    ]),
    row([
      cell("STEP４ 扶養関係", 1900, { shade: true, size: 16 }),
      cell("□ 該当あり（税扶養・健保扶養・手当のいずれか）　□ すべて非該当", 5100, { size: 16 }),
      cell("該当あり → 合算する", 2638, { size: 16 }),
    ]),
    row([
      cell("STEP６ ４要件", 1900, { shade: true, size: 16 }),
      cell("Ａ住居 □充足 □欠落　　Ｂ家計 □充足 □欠落　　Ｃ扶養なし □充足 □欠落　　Ｄ独立収入 □充足 □欠落", 5100, { size: 16 }),
      cell("全て充足 → 別生計／1つでも欠落 → 合算", 2638, { size: 16 }),
    ]),
    row([
      cell("判　定", 1900, { shade: true, size: 16 }),
      cell("□ 合算しない（別生計）　□ 合算する（主宰者：所得割最多の　　　　　　　）　□ 最高階層", 5100, { size: 16 }),
      cell("判定日　　・　　・　　／判定者", 2638, { size: 16 }),
    ]),
  ]
));

// 記入上の注意
children.push(p("【記入にあたっての注意】", { gothic: true, bold: true, size: 18, before: 40, after: 15 }));
children.push(p("・この申立書は、保育料（利用者負担額）の算定において、同居している祖父母等の市町村民税額を合算しないこと（別生計）の判断を受けるために提出するものです。", { size: 16, after: 12 }));
children.push(p("・住民票上世帯分離をしている場合でも、住所（地番）が同一であれば「同居」として扱われ、生活実態が同一生計と認められるときは、祖父母等のうち最多所得の方を家計の主宰者として税額が合算されることがあります。", { size: 16, after: 12 }));
children.push(p("・父母の収入が自治体の定める基準（非課税・年収100〜180万円・生活保護基準など。自治体により異なる）以上の場合は、本申立てによらず父母のみの税額で算定されます。詳細は担当課にお問い合わせください。", { size: 16, after: 8 }));

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
          margin: { top: 1000, bottom: 570, left: 1134, right: 1134 },
        },
      },
      children,
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(process.argv[2] || "form2.docx", buf);
  console.log("written");
});

/**
 * 情報工学基礎 スプレッドシート実習 — データセットアップスクリプト
 * 【高速版】setValues / setFormulas でバッチ処理 → 30秒以内で完了
 * 使い方: 拡張機能 → Apps Script → 貼り付け → setupAll を実行
 */

// ── カラー ──
var H_BG  = '#1a1a1a', H_FG  = '#ffffff';
var F_BG  = '#fff9c4', F_FG  = '#b0b0b0';
var DONE  = '#e8f5e9', WARN  = '#fff3e0', ALERT = '#ffebee';
var SEC   = '#f0f0f0', BOR   = '#d0d0d0';
var RED   = '#c62828', AMBER = '#e65100';

// 日付式（DATE関数ベース）
function d(n) {
  if (n === 0) return '=DATE(2026,5,7)';
  return '=DATE(2026,5,7)' + (n > 0 ? '+' : '') + n;
}

function getOrCreate(ss, name) {
  var sh = ss.getSheetByName(name);
  if (sh) { sh.clearContents(); sh.clearFormats(); }
  else     { sh = ss.insertSheet(name); }
  return sh;
}

function header(sh, row, cols, title, height) {
  var r = sh.getRange(row, 1, 1, cols).merge();
  r.setValue(title).setBackground(H_BG).setFontColor(H_FG)
   .setFontSize(14).setFontWeight('bold').setVerticalAlignment('middle');
  sh.setRowHeight(row, height || 40);
}

function desc(sh, row, cols, text, height) {
  sh.getRange(row, 1, 1, cols).merge()
    .setValue(text).setBackground('#f9f9f9').setFontColor('#555')
    .setFontSize(10).setWrap(true);
  sh.setRowHeight(row, height || 44);
}

function colHeader(sh, row, cols, labels) {
  sh.getRange(row, 1, 1, cols).setValues([labels])
    .setBackground(H_BG).setFontColor(H_FG).setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(row, 32);
}

function borders(sh, row, numRows, numCols) {
  sh.getRange(row, 1, numRows, numCols)
    .setBorder(true, true, true, true, true, true, BOR, SpreadsheetApp.BorderStyle.SOLID);
}

function fmtFormulaCells(sh, row, numRows, col) {
  sh.getRange(row, col, numRows, 1)
    .setBackground(F_BG).setFontColor(F_FG).setFontStyle('italic');
}

// ════════════════════════════════════════
// メイン
// ════════════════════════════════════════
function setupAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSheet_Intro(ss);
  SpreadsheetApp.flush();
  setupSheet_Step1(ss);
  SpreadsheetApp.flush();
  setupSheet_Step2(ss);
  SpreadsheetApp.flush();
  setupSheet_Step3(ss);
  SpreadsheetApp.flush();
  setupSheet_Step4(ss);
  SpreadsheetApp.flush();
  setupSheet_Step5(ss);
  SpreadsheetApp.flush();
  setupSheet_Step6(ss);
  SpreadsheetApp.flush();
  ss.setActiveSheet(ss.getSheetByName('📋 説明'));
  SpreadsheetApp.getUi().alert('✅ セットアップ完了！「📋 説明」から始めてください。');
}

// ════════════════════════════════════════
// 📋 説明シート
// ════════════════════════════════════════
function setupSheet_Intro(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sh = getOrCreate(ss, '📋 説明');
  sh.setTabColor('#1a1a1a');
  [220,160,300,160].forEach(function(w,i){ sh.setColumnWidth(i+1, w); });

  // 全データを一括書き込み
  sh.getRange(1,1,20,4).setValues([
    ['情報工学基礎 — スプレッドシート実習','','',''],
    ['授業日: 2026年5月7日　対象: 1年生　神山まるごと高専','','',''],
    ['','','',''],
    ['━━ 今日の目標 ━━','','',''],
    ['合格','全員到達','プルダウン設定＋ステータスで色が変わる',''],
    ['良',  '目標ライン','COUNTIF/IFで自動集計・自動判定ができる',''],
    ['優',  '発展','グラフ/SPARKLINE/AI関数に挑戦している',''],
    ['','','',''],
    ['━━ シート構成と進め方 ━━','','',''],
    ['シート名','学ぶ関数・機能','内容','目安時間'],
    ['STEP1_予算表','=SUM / =AVERAGE','まるごと祭予算の合計・平均を関数で計算する','15分'],
    ['STEP2_IF判定','=IF / =IFS','期限・ステータスで自動判定テキストを出す','15分'],
    ['STEP3_COUNTIF','=COUNTIF / =COUNTA','完了件数・未着手件数を自動集計する','15分'],
    ['STEP4_タスク管理','プルダウン・条件付き書式','色が自動で変わるタスク管理シートを完成させる','20分'],
    ['STEP5_完成形','全機能の統合','自チームのタスクに書き換えて完成形シートを仕上げる','15分'],
    ['STEP6_発展','SPARKLINE・GAS','発展関数・スクリプトで自動化に挑戦する','発展'],
    ['','','',''],
    ['━━ 関数入力セルの見方 ━━','','',''],
    ['黄色いセル','← ここに関数を入力するセル（ヒントの文字を消して書く）','',''],
    ['グレーの文字','← ヒント。消して自分で書こう','',''],
  ]);

  // タイトル行
  sh.getRange(1,1,1,4).merge().setBackground(H_BG).setFontColor(H_FG)
    .setFontSize(16).setFontWeight('bold').setVerticalAlignment('middle');
  sh.setRowHeight(1, 48);
  sh.getRange(2,1,1,4).merge().setBackground('#333').setFontColor('#ccc').setFontSize(10);

  // セクション見出し（行 4, 9, 16）
  [4,9,18].forEach(function(r){
    sh.getRange(r,1,1,4).merge().setBackground(SEC).setFontWeight('bold').setFontSize(12);
  });

  // 到達目標（行 5-7）
  var lvBg = ['#f1f8e9','#e3f2fd','#f3e5f5'];
  var lvFg = ['#2e7d32','#1565c0','#6a1b9a'];
  [5,6,7].forEach(function(r,i){
    sh.getRange(r,1).setBackground(lvBg[i]).setFontColor(lvFg[i]).setFontWeight('bold');
    sh.getRange(r,1,1,4).setBorder(true,true,true,true,null,null,BOR,SpreadsheetApp.BorderStyle.SOLID);
  });

  // シート構成ヘッダー（行10）
  sh.getRange(10,1,1,4).setBackground(H_BG).setFontColor(H_FG).setFontWeight('bold').setFontSize(11);
  // シート一覧（行11-14）
  sh.getRange(11,1,6,4).setBorder(true,true,true,true,true,true,BOR,SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(11,1,6,1).setFontWeight('bold');
  sh.getRange(15,1,1,4).setBackground('#e0f2f1');  // STEP5 teal
  sh.getRange(16,1,1,4).setBackground('#f3e5f5');  // STEP6 purple

  // 凡例
  sh.getRange(19,1).setBackground(F_BG).setFontWeight('bold');
  sh.getRange(20,1).setFontColor(F_FG);
}

// ════════════════════════════════════════
// STEP1_予算表
// ════════════════════════════════════════
function setupSheet_Step1(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sh = getOrCreate(ss, 'STEP1_予算表');
  sh.setTabColor('#4caf50');
  [200,120,120,120].forEach(function(w,i){ sh.setColumnWidth(i+1, w); });

  // 全データ一括書き込み（13行×4列）
  sh.getRange(1,1,13,4).setValues([
    ['STEP 1 — SUM・AVERAGE で予算表を作ろう','','',''],
    ['● D列（残額）に「=B列-C列」を入れる　● 12行目に =SUM() で合計　● 13行目に =AVERAGE()（発展）','','',''],
    ['','','',''],
    ['','','',''],
    ['','','',''],
    ['品目','予算（円）','支出（円）','残額（円）'],
    ['材料費（チュロス）', 15000, 12000, '← =B7-C7'],
    ['備品費（容器・袋）',  5000,  4800, '← =B8-C8'],
    ['看板・ポスター印刷',  3000,  2500, '← =B9-C9'],
    ['ユニフォーム',        8000,     0, '← =B10-C10'],
    ['その他雑費',          2000,   800, '← =B11-C11'],
    ['合計','← =SUM(B7:B11)','← =SUM(C7:C11)','← =SUM(D7:D11)'],
    ['支出の平均（発展）','','← =AVERAGE(C7:C11)',''],
  ]);

  // 書式まとめて適用
  sh.getRange(1,1,1,4).merge().setBackground(H_BG).setFontColor(H_FG)
    .setFontSize(14).setFontWeight('bold').setVerticalAlignment('middle');
  sh.setRowHeight(1, 40);
  sh.getRange(2,1,1,4).merge().setBackground('#f9f9f9').setFontColor('#555').setFontSize(10).setWrap(true);
  sh.setRowHeight(2, 48);

  sh.getRange(6,1,1,4).setBackground(H_BG).setFontColor(H_FG).setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(6, 32);

  // 数値列まとめて書式
  sh.getRange(7,2,5,2).setNumberFormat('#,##0');  // 予算・支出
  sh.getRange(7,4,5,1).setNumberFormat('#,##0');  // 残額

  // 黄セル（D7:D11 + B12:D12 + C13）
  fmtFormulaCells(sh, 7,  5, 4);   // D7:D11
  sh.getRange(12,2,1,3).setBackground(F_BG).setFontColor(F_FG).setFontStyle('italic').setNumberFormat('#,##0');
  sh.getRange(13,3,1,1).setBackground(F_BG).setFontColor(F_FG).setFontStyle('italic').setNumberFormat('#,##0');

  // 合計行ラベル
  sh.getRange(12,1).setBackground(SEC).setFontWeight('bold');
  // 発展行ラベル
  sh.getRange(13,1).setFontColor('#888').setFontStyle('italic');

  // 枠線まとめて
  borders(sh, 6, 8, 4);
  sh.setFrozenRows(6);
}

// ════════════════════════════════════════
// STEP2_IF判定
// ════════════════════════════════════════
function setupSheet_Step2(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sh = getOrCreate(ss, 'STEP2_IF判定');
  sh.setTabColor('#2196f3');
  [100,210,100,110,160].forEach(function(w,i){ sh.setColumnWidth(i+1, w); });

  // データ（文字列部分）まとめて書き込み
  sh.getRange(1,1,14,5).setValues([
    ['STEP 2 — IF・IFS で自動判定させよう','','','',''],
    ['● E列（判定）に =IF(D列="完了","✅ 完了","❌ 未完了") を入れる　● 期限切れも判定（TODAY()と比較）　● IFSも挑戦（発展）','','','',''],
    ['','','','',''],
    ['','','','',''],
    ['','','','',''],
    ['担当者','タスク名','期限','ステータス','判定（← 関数）'],
    ['田中','保健所に申請書を提出する','','未着手','← ここにIF関数を書く'],
    ['鈴木','試作品を先生に試食してもらう','','進行中','← ここにIF関数を書く'],
    ['山田','チュロスの材料を調達する','','未着手','← ここにIF関数を書く'],
    ['佐藤','看板デザインを完成させる','','完了','← ここにIF関数を書く'],
    ['伊藤','値段設定と原価計算をする','','未着手','← ここにIF関数を書く'],
    ['','','','',''],
    ['💡 ヒント: =IF(D7="完了","✅ 完了", IF(C7<TODAY(),"🚨 期限切れ","🟢 進行中"))','','','',''],
    ['発展: =IFS(D7="完了","✅ 完了", C7<TODAY(),"🚨 期限切れ", C7-TODAY()<=3,"⚠️ まもなく", TRUE,"🟢 余裕あり")','','','',''],
  ]);

  // 期限列（C列）を数式で一括設定
  sh.getRange(7,3,5,1).setFormulas([
    [d(-3)], [d(2)], [d(7)], [d(13)], [d(18)]
  ]);
  sh.getRange(7,3,5,1).setNumberFormat('M/d');

  // 書式まとめて
  sh.getRange(1,1,1,5).merge().setBackground(H_BG).setFontColor(H_FG)
    .setFontSize(14).setFontWeight('bold').setVerticalAlignment('middle');
  sh.setRowHeight(1, 40);
  sh.getRange(2,1,1,5).merge().setBackground('#f9f9f9').setFontColor('#555').setFontSize(10).setWrap(true);
  sh.setRowHeight(2, 52);
  sh.getRange(6,1,1,5).setBackground(H_BG).setFontColor(H_FG).setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(6, 32);

  // 黄セル（E7:E11）
  fmtFormulaCells(sh, 7, 5, 5);

  // ヒント行
  sh.getRange(13,1,1,5).merge().setBackground('#fffde7').setFontColor('#555').setFontSize(10).setFontStyle('italic');
  sh.setRowHeight(13, 36);
  sh.getRange(14,1,1,5).merge().setBackground('#fce4ec').setFontColor('#888').setFontSize(10).setFontStyle('italic');

  borders(sh, 6, 6, 5);
  sh.setFrozenRows(6);
}

// ════════════════════════════════════════
// STEP3_COUNTIF
// ════════════════════════════════════════
function setupSheet_Step3(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sh = getOrCreate(ss, 'STEP3_COUNTIF');
  sh.setTabColor('#ff9800');
  [210,100,100,110,30,30,130,160].forEach(function(w,i){ sh.setColumnWidth(i+1, w); });

  // タスクデータ（A-D列）
  sh.getRange(1,1,15,8).setValues([
    ['STEP 3 — COUNTIF・COUNTA で集計しよう','','','','','','',''],
    ['● H列の集計セルに COUNTIF・COUNTA 関数を入れる　● SPARKLINE で進捗バーを作る（発展）','','','','','','',''],
    ['','','','','','','',''],
    ['','','','','','','',''],
    ['タスク名','担当者','期限','ステータス','','','集計項目','結果（← 関数）'],
    ['保健所に申請書を提出する','田中','','未着手','','','総タスク数','← =COUNTA(A6:A13)'],
    ['試作品を先生に試食してもらう','鈴木','','進行中','','','完了','← =COUNTIF(D6:D13,"完了")'],
    ['チュロスの材料を調達する','山田','','未着手','','','進行中','← =COUNTIF(D6:D13,"進行中")'],
    ['看板デザインを完成させる','佐藤','','完了','','','未着手','← =COUNTIF(D6:D13,"未着手")'],
    ['値段設定と原価計算をする','伊藤','','未着手','','','完了率','← 完了数÷総タスク数'],
    ['当日のシフト表を作る','田中','','未着手','','','進捗バー','← =SPARKLINE(H10,{"charttype","bar";"max",1})（発展）'],
    ['試食会の日程調整をする','鈴木','','進行中','','','',''],
    ['会計報告書のフォーマットを作る','伊藤','','完了','','','',''],
    ['','','','','','','',''],
    ['💡 G・H列に関数を入れよう。ステータスを変えると集計が変わるか確認！','','','','','','',''],
  ]);

  // 期限列（C列, 行6-13）一括設定
  sh.getRange(6,3,8,1).setFormulas([
    [d(-3)],[d(2)],[d(7)],[d(13)],[d(18)],[d(20)],[d(10)],[d(25)]
  ]);
  sh.getRange(6,3,8,1).setNumberFormat('M/d');

  // 書式
  sh.getRange(1,1,1,8).merge().setBackground(H_BG).setFontColor(H_FG)
    .setFontSize(14).setFontWeight('bold').setVerticalAlignment('middle');
  sh.setRowHeight(1, 40);
  sh.getRange(2,1,1,8).merge().setBackground('#f9f9f9').setFontColor('#555').setFontSize(10).setWrap(true);
  sh.setRowHeight(2, 40);

  // ヘッダー行5
  sh.getRange(5,1,1,4).setBackground(H_BG).setFontColor(H_FG).setFontWeight('bold').setFontSize(11);
  sh.getRange(5,7,1,2).setBackground(H_BG).setFontColor(H_FG).setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(5, 32);

  // 集計黄セル（H6:H11）
  fmtFormulaCells(sh, 6, 6, 8);
  sh.getRange(10,8).setNumberFormat('0%'); // 完了率

  // 完了行に薄緑
  sh.getRange(9,1,1,4).setBackground(DONE);   // 看板（完了）
  sh.getRange(13,1,1,4).setBackground(DONE);  // 会計報告（完了）

  // 枠線
  borders(sh, 5, 9, 4);
  sh.getRange(5,7,7,2).setBorder(true,true,true,true,true,true,BOR,SpreadsheetApp.BorderStyle.SOLID);

  // ヒント行
  sh.getRange(15,1,1,8).merge().setBackground('#fffde7').setFontColor('#555').setFontSize(10);

  sh.setFrozenRows(5);
}

// ════════════════════════════════════════
// STEP4_タスク管理
// ════════════════════════════════════════
function setupSheet_Step4(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sh = getOrCreate(ss, 'STEP4_タスク管理');
  sh.setTabColor('#9c27b0');
  [28,220,90,90,110].forEach(function(w,i){ sh.setColumnWidth(i+1, w); });

  var HEADER_ROW = 9;

  // 説明・ルール一括書き込み（8行）
  sh.getRange(1,1,8,5).setValues([
    ['STEP 4 — 完成形：プルダウン＋条件付き書式','','','',''],
    ['① E列（ステータス）にプルダウンを設定 [データ→データの入力規則→リスト→未着手,進行中,完了]\n② 条件付き書式で完了→薄緑、期限切れ→薄赤、3日以内→薄黄\n③ 自分のチームのタスクに書き換えて提出する','','','',''],
    ['','','','',''],
    ['▼ 条件付き書式の設定（書式 → 条件付き書式 → カスタム数式）','','','',''],
    ['条件①', '=$E5="完了"', '→ 書式: テキスト: グレー、背景: 薄緑','',''],
    ['条件②', '=AND($E5<>"完了",$C5<TODAY())', '→ 書式: 背景: 薄赤（期限切れ）','',''],
    ['条件③', '=AND($E5<>"完了",$C5-TODAY()<=3,$C5>=TODAY())', '→ 書式: 背景: 薄黄（まもなく）','',''],
    ['','','','',''],
  ]);

  // タスクデータ（10件）一括書き込み
  var tasks = [
    ['保健所に申請書を提出する',       '田中', '未着手', -3],
    ['試作品（チュロス）を先生に試食',  '鈴木', '進行中',  2],
    ['チュロスの材料を業者に発注',      '山田', '未着手',  7],
    ['看板・メニューボードをデザイン',  '佐藤', '完了',   13],
    ['値段設定と原価計算をする',        '伊藤', '未着手', 18],
    ['当日のシフト表を作る',            '田中', '未着手', 20],
    ['試食会の日程調整',                '鈴木', '進行中', 10],
    ['会計報告書のフォーマット作成',    '伊藤', '完了',   25],
    ['当日の販売マニュアルを作る',      '山田', '未着手', 22],
    ['材料の保管場所を決める',          '佐藤', '未着手',  5],
  ];

  // ヘッダー行
  sh.getRange(HEADER_ROW,1,1,5).setValues([['#','タスク名','期限','担当者','ステータス ▼']]);
  sh.getRange(HEADER_ROW,1,1,5).setBackground(H_BG).setFontColor(H_FG).setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(HEADER_ROW, 32);

  // タスク番号・タスク名・担当者・ステータスを一括書き込み
  var taskValues = tasks.map(function(t,i){ return [i+1, t[0], '', t[1], t[2]]; });
  sh.getRange(HEADER_ROW+1,1,10,5).setValues(taskValues);

  // 期限列（C列）を一括で数式設定
  sh.getRange(HEADER_ROW+1,3,10,1).setFormulas(
    tasks.map(function(t){ return [d(t[3])]; })
  );
  sh.getRange(HEADER_ROW+1,3,10,1).setNumberFormat('M/d');

  // 書式
  sh.getRange(1,1,1,5).merge().setBackground(H_BG).setFontColor(H_FG)
    .setFontSize(14).setFontWeight('bold').setVerticalAlignment('middle');
  sh.setRowHeight(1, 40);
  sh.getRange(2,1,1,5).merge().setBackground('#f9f9f9').setFontColor('#555')
    .setFontSize(10).setWrap(true);
  sh.setRowHeight(2, 72);

  sh.getRange(4,1,1,5).merge().setBackground('#ede7f6').setFontWeight('bold').setFontSize(11);

  // 条件付き書式ルール行（5-7）
  sh.getRange(5,1,3,1).setBackground('#f3e5f5').setFontWeight('bold');
  sh.getRange(5,2,3,1).setBackground(F_BG).setFontColor(F_FG).setFontFamily('Courier New').setFontSize(10);
  sh.getRange(5,3,1,3).merge();  // 1行ずつマージ
  sh.getRange(6,3,1,3).merge();
  sh.getRange(7,3,1,3).merge();
  sh.getRange(5,3).setFontColor('#666').setFontSize(10);
  sh.getRange(6,3).setFontColor('#666').setFontSize(10);
  sh.getRange(7,3).setFontColor('#666').setFontSize(10);
  sh.getRange(5,1,3,5).setBorder(true,true,true,true,true,true,BOR,SpreadsheetApp.BorderStyle.SOLID);

  // タスク行の色付け（参考：生徒が条件付き書式で再現する）
  tasks.forEach(function(t, i) {
    var row = HEADER_ROW + 1 + i;
    if (t[2] === '完了') {
      sh.getRange(row,1,1,5).setFontColor('#999');
    } else if (t[3] < 0) {
      sh.getRange(row,3).setBackground(ALERT).setFontColor(RED).setFontWeight('bold');
    } else if (t[3] <= 3) {
      sh.getRange(row,3).setBackground(WARN).setFontColor(AMBER).setFontWeight('bold');
    }
  });

  // 番号列の文字色
  sh.getRange(HEADER_ROW+1,1,10,1).setFontColor('#ccc').setHorizontalAlignment('center');

  // 枠線
  borders(sh, HEADER_ROW, 11, 5);

  // 追加行（空行）
  var addRow = HEADER_ROW + 11;
  sh.getRange(addRow,1,1,5).merge()
    .setValue('↓ ここから自分のチームのタスクを追加しよう（既存タスクを書き換えてもOK）')
    .setBackground('#fffde7').setFontColor('#888').setFontStyle('italic').setFontSize(10);

  var emptyNums = [];
  for (var i = 0; i < 8; i++) emptyNums.push([11+i,'','','','']);
  sh.getRange(addRow+1,1,8,5).setValues(emptyNums);
  sh.getRange(addRow+1,1,8,5).setBackground('#fafafa');
  sh.getRange(addRow+1,1,8,1).setFontColor('#ddd').setHorizontalAlignment('center');
  sh.getRange(addRow+1,1,8,5).setBorder(true,true,true,true,true,true,BOR,SpreadsheetApp.BorderStyle.SOLID);

  // プルダウン（E列）
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未着手','進行中','完了'], true)
    .setAllowInvalid(false).build();
  sh.getRange(HEADER_ROW+1,5,20,1).setDataValidation(rule);

  sh.setFrozenRows(HEADER_ROW);
}

// ════════════════════════════════════════
// STEP5_完成形
// ════════════════════════════════════════
function setupSheet_Step5(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sh = getOrCreate(ss, 'STEP5_完成形');
  sh.setTabColor('#00bcd4');
  [28,200,80,80,100,30,140,120].forEach(function(w,i){ sh.setColumnWidth(i+1, w); });

  var TASK_ROW = 12;
  var TASK_START = 13;
  var STUDENT_ROW = 23;

  // 上部データ一括書き込み
  sh.getRange(1,1,11,8).setValues([
    ['STEP 5 — 完成形：タスク管理シートを仕上げよう','','','','','','',''],
    ['STEP 1〜4を組み合わせた完成形。自分のチームのタスクに書き換えてカスタマイズしよう。','','','','','','',''],
    ['','','','','','','',''],
    ['▼ 最終課題チェックリスト','','','','','','📊 自動集計（数式は変更不要）',''],
    ['達成項目','','','','','','集計項目','結果（自動）'],
    ['① 自チームのタスクを5件以上入力する（タスク名・担当者・期限）','','','','','','総タスク数',''],
    ['② E列のプルダウンでステータスを「未着手 / 進行中 / 完了」に設定する','','','','','','完了',''],
    ['③ ステータスを変えると右の集計が自動で変わることを確認する','','','','','','進行中',''],
    ['④ 書式 → 条件付き書式で期限切れ・完了の色分けを設定する','','','','','','未着手',''],
    ['⑤ 完了率（完了数 ÷ 総タスク数）を確認する','','','','','','完了率',''],
    ['⑥ 【発展】SPARKLINE 関数で進捗バーをH列に作る','','','','','','進捗バー',''],
  ]);

  // タスクヘッダー
  sh.getRange(TASK_ROW,1,1,5).setValues([['#','タスク名','期限','担当者','ステータス ▼']]);

  // タスクデータ
  var tasks = [
    ['保健所に申請書を提出する',       '田中', '未着手', -3],
    ['試作品（チュロス）を先生に試食',  '鈴木', '進行中',  2],
    ['チュロスの材料を業者に発注',      '山田', '未着手',  7],
    ['看板・メニューボードをデザイン',  '佐藤', '完了',   13],
    ['値段設定と原価計算をする',        '伊藤', '未着手', 18],
    ['当日のシフト表を作る',            '田中', '未着手', 20],
    ['試食会の日程調整',                '鈴木', '進行中', 10],
    ['会計報告書のフォーマット作成',    '伊藤', '完了',   25],
    ['当日の販売マニュアルを作る',      '山田', '未着手', 22],
    ['材料の保管場所を決める',          '佐藤', '未着手',  5],
  ];
  var taskVals = tasks.map(function(t,i){ return [i+1, t[0], '', t[1], t[2]]; });
  sh.getRange(TASK_START,1,10,5).setValues(taskVals);
  sh.getRange(TASK_START,3,10,1).setFormulas(tasks.map(function(t){ return [d(t[3])]; }));
  sh.getRange(TASK_START,3,10,1).setNumberFormat('M/d');

  // 進捗数式（H列 行6-11）— 入力済みとして提示
  sh.getRange(6,8,6,1).setFormulas([
    ['=COUNTA(B13:B33)'],
    ['=COUNTIF(E13:E33,"完了")'],
    ['=COUNTIF(E13:E33,"進行中")'],
    ['=COUNTIF(E13:E33,"未着手")'],
    ['=IFERROR(H7/H6,0)'],
    ['=SPARKLINE(H10,{"charttype","bar";"max",1})'],
  ]);
  sh.getRange(10,8).setNumberFormat('0%');

  // 書式
  sh.getRange(1,1,1,8).merge().setBackground(H_BG).setFontColor(H_FG)
    .setFontSize(14).setFontWeight('bold').setVerticalAlignment('middle');
  sh.setRowHeight(1, 40);
  sh.getRange(2,1,1,8).merge().setBackground('#f9f9f9').setFontColor('#555').setFontSize(10).setWrap(true);
  sh.setRowHeight(2, 40);

  sh.getRange(4,1,1,5).merge().setBackground('#e0f2f1').setFontWeight('bold').setFontSize(11);
  sh.getRange(4,7,1,2).merge().setBackground(H_BG).setFontColor(H_FG).setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(4, 32);
  sh.getRange(5,1,1,5).merge().setBackground('#b2dfdb').setFontColor('#555').setFontWeight('bold').setFontSize(10);
  sh.getRange(5,7,1,2).setBackground(H_BG).setFontColor(H_FG).setFontWeight('bold').setFontSize(10);
  sh.setRowHeight(5, 28);

  for (var r = 6; r <= 11; r++) {
    sh.getRange(r,1,1,5).merge().setFontSize(10);
    if (r % 2 === 0) sh.getRange(r,1,1,5).setBackground('#f9fafb');
  }
  sh.getRange(6,7,6,1).setFontWeight('bold').setFontSize(10);
  sh.getRange(6,7,6,2).setBorder(true,true,true,true,true,true,BOR,SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(TASK_ROW,1,1,5).setBackground(H_BG).setFontColor(H_FG).setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(TASK_ROW, 32);

  tasks.forEach(function(t,i) {
    var row = TASK_START + i;
    if (t[2] === '完了')  sh.getRange(row,1,1,5).setFontColor('#999');
    else if (t[3] < 0)   sh.getRange(row,3).setBackground(ALERT).setFontColor(RED).setFontWeight('bold');
    else if (t[3] <= 3)  sh.getRange(row,3).setBackground(WARN).setFontColor(AMBER).setFontWeight('bold');
  });
  sh.getRange(TASK_START,1,10,1).setFontColor('#ccc').setHorizontalAlignment('center');
  sh.getRange(TASK_ROW,1,11,5).setBorder(true,true,true,true,true,true,BOR,SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(STUDENT_ROW,1,1,5).merge()
    .setValue('↓ ここから自分のチームのタスクに書き換えよう（タスク名・担当者・期限・ステータスを変更）')
    .setBackground('#e0f7fa').setFontColor('#00695c').setFontStyle('italic').setFontSize(10);
  sh.setRowHeight(STUDENT_ROW, 36);

  var emptyRows = [];
  for (var j = 0; j < 10; j++) emptyRows.push([11+j,'','','','']);
  sh.getRange(STUDENT_ROW+1,1,10,5).setValues(emptyRows);
  sh.getRange(STUDENT_ROW+1,1,10,5).setBackground('#fafafa');
  sh.getRange(STUDENT_ROW+1,1,10,1).setFontColor('#ddd').setHorizontalAlignment('center');
  sh.getRange(STUDENT_ROW+1,1,10,5).setBorder(true,true,true,true,true,true,BOR,SpreadsheetApp.BorderStyle.SOLID);

  var rule5 = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未着手','進行中','完了'], true)
    .setAllowInvalid(false).build();
  sh.getRange(TASK_START,5,21,1).setDataValidation(rule5);

  sh.setFrozenRows(TASK_ROW);
}

// ════════════════════════════════════════
// STEP6_発展
// ════════════════════════════════════════
function setupSheet_Step6(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sh = getOrCreate(ss, 'STEP6_発展');
  sh.setTabColor('#e91e63');
  [30,160,270,180].forEach(function(w,i){ sh.setColumnWidth(i+1, w); });

  // 発展関数一覧
  sh.getRange(1,1,2,4).setValues([
    ['STEP 6 — 発展：さらに使いこなす','','',''],
    ['基本をマスターしたら「自動化」へ。発展関数と GAS でスプレッドシートをさらに進化させよう。','','',''],
  ]);
  sh.getRange(4,1,1,4).setValues([['📊 発展関数一覧','','','']]);
  sh.getRange(5,1,1,4).setValues([['#','機能','数式例','できること']]);
  sh.getRange(6,1,6,4).setValues([
    [1,'SPARKLINE','=SPARKLINE(C2/B2,{"charttype","bar";"max",1})','セル内に棒グラフを表示。完了率の可視化に最適'],
    [2,'FILTER','=FILTER(A2:E20,E2:E20<>"完了")','条件を満たす行を自動抽出。未完了タスク一覧に便利'],
    [3,'QUERY','=QUERY(A:E,"SELECT * WHERE E=\'未着手\' ORDER BY D")','SQL風でデータ絞り込み・並び替え'],
    [4,'IMPORTRANGE','=IMPORTRANGE("スプシURL","シート名!A:E")','別ファイルのデータをリアルタイム取得'],
    [5,'COUNTIFS','=COUNTIFS(D:D,"田中",E:E,"完了")','複数条件でカウント（担当者 かつ 完了）'],
    [6,'AI関数','=AI("チュロス屋の準備タスクを5件書いて")','Gemini連携でAIが直接回答（Workspace限定）'],
  ]);

  // SPARKLINEデモ
  sh.getRange(13,1,1,4).setValues([['📈 SPARKLINE デモ（担当者別完了率）','','','']]);
  sh.getRange(14,1,1,4).setValues([['担当者','タスク数','完了数','進捗バー（=SPARKLINE）']]);
  sh.getRange(15,1,5,3).setValues([
    ['田中',5,3],['鈴木',4,4],['山田',6,2],['佐藤',3,3],['伊藤',4,1],
  ]);
  sh.getRange(15,4,5,1).setFormulas([
    ['=SPARKLINE(C15/B15,{"charttype","bar";"max",1;"color1","#16a34a"})'],
    ['=SPARKLINE(C16/B16,{"charttype","bar";"max",1;"color1","#2196f3"})'],
    ['=SPARKLINE(C17/B17,{"charttype","bar";"max",1;"color1","#ff9800"})'],
    ['=SPARKLINE(C18/B18,{"charttype","bar";"max",1;"color1","#16a34a"})'],
    ['=SPARKLINE(C19/B19,{"charttype","bar";"max",1;"color1","#ff9800"})'],
  ]);

  // GASセクション
  sh.getRange(21,1,1,4).setValues([['🤖 GAS（Google Apps Script）でさらに自動化','','','']]);
  sh.getRange(22,1,1,4).setValues([['拡張機能 → Apps Script を開いてJavaScriptで書く。授業サイトのSTEP 6タブにコードあり。','','','']]);
  sh.getRange(23,1,1,4).setValues([['テーマ','機能','実行タイミング','難易度']]);
  sh.getRange(24,1,6,4).setValues([
    ['📧 期限アラートメール','期限3日以内のタスクを自動でメール通知','毎朝9時（時間トリガー）','★☆☆'],
    ['🔔 onEdit 完了日記録','「完了」変更時に隣のセルへ日時を自動記録','セル編集時（onEdit）','★☆☆'],
    ['📝 フォーム連携','フォーム送信でタスクを自動追加','フォーム送信時','★★☆'],
    ['📦 完了アーカイブ','完了タスクを別シートへ自動移動','手動実行','★★☆'],
    ['📊 週次レポート','担当者別完了率を週1回メール','毎週月曜（時間トリガー）','★★☆'],
    ['💬 Slack通知','期限切れをSlack/Discordへ自動投稿','時間トリガー','★★★'],
  ]);

  // 書式
  sh.getRange(1,1,1,4).merge().setBackground('#ad1457').setFontColor(H_FG)
    .setFontSize(14).setFontWeight('bold').setVerticalAlignment('middle');
  sh.setRowHeight(1, 40);
  sh.getRange(2,1,1,4).merge().setBackground('#fce4ec').setFontColor('#555').setFontSize(10).setWrap(true);
  sh.setRowHeight(2, 40);

  sh.getRange(4,1,1,4).merge().setBackground(H_BG).setFontColor(H_FG).setFontWeight('bold').setFontSize(12);
  sh.setRowHeight(4, 36);
  sh.getRange(5,1,1,4).setBackground(H_BG).setFontColor(H_FG).setFontWeight('bold').setFontSize(10);
  sh.setRowHeight(5, 28);
  sh.getRange(6,1,6,4).setBorder(true,true,true,true,true,true,BOR,SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(6,1,6,1).setHorizontalAlignment('center').setFontColor('#aaa');
  sh.getRange(6,3,6,1).setBackground(F_BG).setFontColor(F_FG).setFontFamily('Courier New').setFontSize(10).setWrap(true);
  [7,9,11].forEach(function(r){ sh.getRange(r,1,1,4).setBackground('#fafafa'); });

  sh.getRange(13,1,1,4).merge().setBackground(H_BG).setFontColor(H_FG).setFontWeight('bold').setFontSize(12);
  sh.setRowHeight(13, 36);
  sh.getRange(14,1,1,4).setBackground(H_BG).setFontColor(H_FG).setFontWeight('bold').setFontSize(10);
  sh.setRowHeight(14, 28);
  sh.getRange(15,1,5,4).setBorder(true,true,true,true,true,true,BOR,SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeights(15, 5, 32);

  sh.getRange(21,1,1,4).merge().setBackground(H_BG).setFontColor(H_FG).setFontWeight('bold').setFontSize(12);
  sh.setRowHeight(21, 36);
  sh.getRange(22,1,1,4).merge().setBackground('#f9f9f9').setFontColor('#555').setFontSize(10).setWrap(true);
  sh.setRowHeight(22, 40);
  sh.getRange(23,1,1,4).setBackground(H_BG).setFontColor(H_FG).setFontWeight('bold').setFontSize(10);
  sh.setRowHeight(23, 28);
  sh.getRange(24,1,6,4).setBorder(true,true,true,true,true,true,BOR,SpreadsheetApp.BorderStyle.SOLID);
  [25,27,29].forEach(function(r){ sh.getRange(r,1,1,4).setBackground('#fafafa'); });

  var diffBg = ['#e8f5e9','#e8f5e9','#fff3e0','#fff3e0','#fff3e0','#ffebee'];
  for (var i = 0; i < 6; i++) {
    sh.getRange(24+i,4).setBackground(diffBg[i]).setFontWeight('bold').setHorizontalAlignment('center');
  }

  sh.getRange(31,1,1,4).merge()
    .setValue('🎓 GASはJavaScriptで書く。今日学んだIF・COUNTIFの考え方はそのまま活かせる。次はWebアプリ開発（HTML/CSS/React）へ！')
    .setBackground(H_BG).setFontColor('#4ade80').setFontSize(11).setFontWeight('bold').setWrap(true);
  sh.setRowHeight(31, 48);
}

# Claude Cowork用 EZPost自動入力スキル

`ezpost-entry/` は、デスクトップアプリ **Claude Cowork** に追加して使うスキルです。
管理者ページの発送タブでコピーした「EZPost入力データ」を貼り付けて
「EZPost入力して」と頼むと、中華郵政EZPostの発送フォーム入力を自動化できます。

## インストール方法

1. `ezpost-entry` フォルダ(中に `SKILL.md`)をパソコンにコピーする
2. Claude Cowork の設定からスキル(Skills)として `ezpost-entry` フォルダを追加する
   - スキル追加のUIがない場合は、`~/.claude/skills/ezpost-entry/SKILL.md` に配置すると
     Claude Code / Cowork から利用できます
3. Cowork のチャットで「EZPost入力して」と入力し、続けて
   管理者ページでコピーした「EZPost入力データ」を貼り付ける

## 使い方の流れ

1. 管理者ページ「発送」タブ → 発送待ちの注文で「📋 EZPost入力データをコピー」
2. Cowork に貼り付けて「EZPost入力して」
3. Cowork がブラウザで https://ezpost.post.gov.tw/ を開く(ログインは自分で行う)
4. 宛先・内容品・重量が自動入力される
5. **送信前に必ず内容を確認して承認**(承認するまで送信されません)
6. 発行された追跡番号を管理者ページの「追跡番号」欄に入力 → 「✈️ 発送済にする」

## 安全設計

- ログインID・パスワードはCoworkに渡さない(ログインは人間が行う)
- 発送の確定(送信)は必ず人間の承認後
- サイトの画面が変わった場合は自動操作を中断して報告する

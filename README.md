# Excel_FGT

FortiGateのコンフィグ（`show full-configuration` / バックアップファイル）から、
アドレス・ポリシーなどのセクションをExcelのシートに一覧化するVBAマクロ。

## 使い方

1. 新しいブックを `.xlsm` で保存
2. `Alt+F11` → 挿入 → 標準モジュール → `FgtConfigExtract.bas` の中身を丸ごと貼り付け
3. `RunFgtExtract` を実行 → コンフィグファイルを選択
4. 初回は「設定」シートが自動作成される
   - A列: 抽出するセクション名（`firewall policy` など。自由に追加・削除可）
   - B列: 抽出件数（自動）
   - D列: 出力しないキー（`uuid`、`passwd`、`psksecret` などは最初から除外）

## 仕様

- `config / edit / set / next / end` を汎用的に解析。列は登場した `set` のキーから自動生成
- 複数値（`set srcaddr "A" "B"`）はセル内改行で連結（`set ip 10.0.0.1 255.255.255.0` のような引用符なしは空白区切り）
- VDOM構成に対応（VDOM列を自動追加）
- `edit` を持たないセクション（`system global`、`router bgp` 等）は1行で出力
- ネストした `config` は `親config名.キー` の列にまとめる
- 前回の結果シートは実行時に上書き（該当なしのセクションはクリア）

## 動作確認用

`sample/sample_fgt.conf` はダミーのコンフィグ。実機のコンフィグはリポジトリに置かないこと。

## 注意

- TeraTermログから読む場合は `config system console` → `set output standard` で
  `--More--` を出さずに取得するのがおすすめ（残っていても除去はする）
- 文字コードはUTF-8前提

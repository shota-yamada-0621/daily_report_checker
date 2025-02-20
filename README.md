#### 日報フォーマットチェックツール

日報の入力ミスを見つけるツールです。
対象年度のシート切り出しや、全シートのアクティブセルをA1に移動するコマンドもあります。

目視確認しなければいけないところ＝C6セルに名前が入っているかどうか


#### python3.12 のインストールが必要です

https://blog.pyq.jp/entry/python_install_231102_win

#### poetry のインストール

```
py -3.12 -m pip install poetry
py -3.12 -m poetry config virtualenvs.in-project true
py -3.12 -m poetry install
```

#### 実行コマンド

日報の内容チェック

```
py -3.12 -m poetry run python app/main.py check 202404 202502
```
**(チェックできる内容)**

```
シート名のチェック
第一引数（開始年月）から第二引数（終了年月）までの範囲で、各週の月曜日始まり・日曜日終わりのシート名リストを生成する
生成したシート名リストと、実際のExcelファイル内のシート名を比較し、一致しているかを確認する
存在すべきシートが欠落している場合、警告を出す
不適切なシート名（期待されるフォーマット YYYYMMDD_YYYYMMDD に合致しないもの）が存在する場合、警告を出す
シート内のセル値チェック
H列（H10〜H16）の日付・数式チェック
シート名の形式が YYYYMMDD_YYYYMMDD であることを確認する
H10の値がシート名の開始日と一致しているか確認する
H11〜H16 の値が =H10+1 のように、H10を基準に1日ずつ増加する数式になっているか確認する
H10～H16 に対応する A列（A10, A16, A22, A28, A34, A40, A46）の数式をチェック
A10: =MONTH(H10)
A11: =DAY(H10)
A12: ="("&TEXT(H10, "aaa")&")"
B列の固定値チェック
以下のセルに指定の固定値が設定されているか確認する：
B10, B16, B22, B28, B34, B40, B46: "月"
B11, B17, B23, B29, B35, B41, B47: "日"
C列（業務内容）の入力チェック
C9, C15, C21, C27, C33, C39, C45 が未入力でないかチェックする
C39, C45 の値が "休暇" であることをチェックする
A57, F4, C6の特定セルチェック
A57: =H14 という数式が入力されているか
F4: "miracleave株式会社" という文字列が入力されているか
C6: 未入力でないか（数値・文字列を含む）
祝日・土日のチェック
H10〜H16 の日付が祝日または土日の場合、対応する C列 のセルに "祝日" または "休暇" が入力されているかチェックする
平日の場合、C列 に "祝日" または "休暇" が入力されていないかチェックする
```




対象期間のシート切り出し(少しだけ時間掛かります)

```
py -3.12 -m poetry run python app/main.py cut 202404 202502
```

#### 機能について

- どの機能も自身の日報フォーマットの Excel 自身を target ディレクトリに格納しておく必要があります。

- 基本的に操作する対象の日報ファイルは 1 つです。

#### バッチ実行コマンド

実行テスト用

```
py -3.11 -m poetry run python app/main.py test
```

シート名チェック

```
py -3.11 -m poetry run python app/main.py sheet-name-check
```

## 開発環境

| 言語・フレームワーク | バージョン |
| -------------------- | ---------- |
| Docker               | 24.0.6     |
| Docker Compose       | 2.21.0     |
| Nginx                | 1.25.5     |
| Spring Boot          | 3.2.5      |
| Gradle               | 8.7.0      |
| Java                 | 17         |
| MySQL                | 8.0.36     |
| Node.js              | 20.12.2    |
| npm                  | 10.5.0     |
| Vite                 | 5.2.3      |
| Vue.js               | 3.4.21     |

## ディレクトリ構成

<pre>
.
├── docker                               Docker開発環境
│   ├── api                             Spring Boot
│   │   └── Dockerfile
│   ├── app                             Vue.js
│   │   └── Dockerfile
│   ├── db                              MySQL
│   │   ├── initdb.d
│   │   │   └── initdb.sql            DB初期データ
│   │   ├── Dockerfile
│   │   └── my.cnf                     設定ファイル
│   ├── web                             Nginx
│   │   ├── default.conf               設定ファイル
│   │   └── Dockerfile
│   ├── .env                            環境変数関連
│   └── compose.yml                     Docker Compose設定
├── document                             ドキュメント
│   ├── api                             API設計書
│   └── db                              データベース関連
│       ├── ddl                         テーブルDDL
│       └── table                       テーブル定義書
└── source                               ソースコード
    ├── api                              API（Spring Boot）
    └── app                              Webアプリ（Vue.js）
</pre>

## 環境変数一覧

| 変数名              | 役割                           | デフォルト値 |
| ------------------- | ------------------------------ | ------------ |
| APP_MODE            | アプリケーションモード         | dev          |
| APP_LOG_LEVEL       | アプリケーションログ出力レベル | debug        |
| MYSQL_ROOT_PASSWORD | MySQL のルートパスワード       | root         |
| MYSQL_DATABASE      | MySQL のデータベース名         | miraedge     |
| MYSQL_USER          | MySQL のユーザ名               | mira         |
| MYSQL_PASSWORD      | MySQL のパスワード             | mira         |

### 機能実装時

(機能実装やソース修正時に必要な各コマンド)  
※コミットメッセージのルールも参照をお願いします

```
$ git flow feature start <ブランチ名>
$ git add <ファイル名>
$ git commit -m "<メッセージ>"
$ git push -u origin feature/<ブランチ名>
$ git flow feature finish <feature name>
$ git push -u origin develop
```

### リリース時

(develop から master ブランチへの反映と ver タグ付けを行う)

```
$ git flow release start <バージョン>
$ git flow release finish <バージョン>
$ git push -u origin develop
$ git push -u origin master
$ git push --tag
```

### 本番環境で修正が必要な場合

(本番用ブランチである master で修正が必要な場合)

```
$ git flow hotfix start <バージョン>
（ソース修正作業）
$ git flow hotfix finish <バージョン>
$ git push -u origin develop
$ git push -u origin master
$ git push --tag
```

### コミットメッセージのルール

下記方式で作成してください。  
"plefix: (日本語で修正内容を分かりやすく記述)"

(参考) https://qiita.com/konatsu_p/items/dfe199ebe3a7d2010b3e

#### 使用する plefix 一覧

```
feat: 新しい機能
fix: バグの修正
docs: ドキュメントのみの変更
style: 空白、フォーマット、セミコロン追加など
refactor: 仕様に影響がないコード改善(リファクタ)
perf: パフォーマンス向上関連
test: テスト関連
chore: ビルド、補助ツール、ライブラリ関連
```

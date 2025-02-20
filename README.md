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
py -3.11 -m poetry run python app/main.py 202404 202502
```


#### 機能について

- どの機能も自身の日報フォーマットの Excel 自身を target ディレクトリに格納しておく必要があります。

- 基本的に操作する対象の日報ファイルは 1 つです。

- シート名チェック機能

対象の Excel オブジェクトが持つシート名の書式と整合性が正しいか確認します。
書式と整合性とは、シート名に関して以下を満たすことを指します。

- シート名が YYYYMMDD_YYYYMMDD となっていること
- 上記 YYYYMMDD の左辺が月曜日、右辺が日曜日であること
- 日付が連番になっていること

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

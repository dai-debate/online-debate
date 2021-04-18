# オンラインディベート運営支援ツール

## 概要

本ツールは、[Zoom](https://zoom.us/) をプラットフォームとして使用するオンラインディベートの運営を支援するための各種機能を実装したものです。
対戦スケジュールの管理や、ジャッジの投票や採点の入力、集計結果の表示等には [Google Spreadsheets](https://www.google.com/intl/ja_jp/sheets/about/) を使用します。

現時点で実装されている機能は以下の通りです。

* 対戦スケジュール表を基に、対応するZoomミーティングを生成する
* 対戦スケジュール表を基に、投票・採点記入用シートを生成する
* 各々の試合用の投票・採点記入用シートから結果を集計するリンクを生成し、一覧化する

また、本ツールは、[日本ディベート協会](https://japan-debate-association.org/)主催のJDA大会、および[全国教室ディベート連盟](https://nade.jp/)主催のディベート甲子園での使用を念頭に置いて作成されていますが、テンプレートと設定ファイルを修正することにより、他の試合形式の大会での使用もできるように設計されています。

## 要求事項

本ツールを使用するには以下のものが必要となります。

* PC (Windows 10 64bitで動作を確認していますが、macやLinuxでも動作できるはずです)
* [Git](https://git-scm.com/) または [GitHub Desktop](https://desktop.github.com/) (ツール一式の取得にZipアーカイブ形式を用いる場合は不要です)
* テキストエディタ ([Visual Studio Code](https://code.visualstudio.com/)を推奨)
* [Python](https://www.python.org/) (ver. 3.8.6 for Windowsで動作を確認)
* [Google アカウント](https://www.google.com/intl/ja/account/about/)
* [Zoom アカウント](https://zoom.us/signup)
  * 本ツールの動作確認目的であれば、無償の基本ライセンスで構いませんが、実際の大会を行う場合はミーティングの時間制限を避けるため、プロ以上のライセンスを並行試合数分取得する必要があります。

## セットアップ手順

本ツールを使用可能にするまでに、最初に必要な手順を以下に示します。

* [オンラインディベート運営支援ツール一式の取得](docs/get-scripts.md)
* [Python および依存パッケージのインストール](docs/install-python.md)
* [Google API の設定](docs/google-api.md)
* [Zoom API の設定](docs/zoom-api.md)
* [Zoom ユーザーの作成](docs/zoom-create-user.md)

## 使用方法

以下に本ツールの代表的な使い方を示します。
対戦スケジュール表や設定ファイルの書き方など、より詳細な使い方は「[本ツールの詳細な使い方](docs/how-to-use.md)」を参照して下さい。

### Zoomミーティングの生成

対戦スケジュール表と設定ファイルを編集した上で、コマンドプロンプトを開き以下のコマンドを実行します。

```console
python manage.py -c config.yaml generate-room
```

### 投票・採点記入用シートの生成

対戦スケジュール表と設定ファイルを編集した上で、コマンドプロンプトを開き以下のコマンドを実行します。

```console
python manage.py -c config.yaml generate-ballot
```

## 制限事項

* ドキュメントの一部が準備中です。随時更新していきます。
* ミーティングやシートの生成は本ツールで一括で行えますが、修正や削除は現時点で未実装です。順次対応していきます。

## 開発者

小山 雄輔

※本ツールの開発は、開発者が所属するいかなる組織の責任下で行われたものではなく、個人的に行われたものです。

## ライセンス

本ツールは、[MITライセンス](https://opensource.org/licenses/mit-license.php) の下に公開しており、無償利用/商用利用を問わず、複製・再配布・派生物の作成等をすることができます。
ただし、本ツールを用いたことによる損害等については責任を負わないものとします。

詳細は、[LICENSE.md](LICENSE.md) を参照してください。

## 改善のための貢献

* 使い方に関する質問や不具合の報告、改善の提案は [Issues](https://github.com/dai-debate/online-debate/issues) に書き込んで下さい。
  可能な範囲で対応します。
* Python のプログラミングに詳しい方は、具体的修正提案を [Pull request](https://github.com/dai-debate/online-debate/pulls) として投稿して下さるのも歓迎です。

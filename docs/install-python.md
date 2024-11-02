# Python および依存パッケージのインストール

## uv のインストール

本ツールは、Python および依存パッケージのインストールに、パッケージマネージャー [uv](https://github.com/astral-sh/uv) を使用します。
まず始めに uv のインストールを行ってください。

### Windows PC の場合

* Windows ターミナル、コマンドプロンプト等のコンソールアプリを立ち上げ、以下のコマンドを実行します。

  ```shell
  powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
  ```

  * `powershell` が見つからないというようなエラーになる場合は、
    `powershell` の部分を `pwsh` に置き換えてみてください。

### Linux (WLS2を含む) の場合

* シェルから以下ののコマンドを実行します。

  ```shell
  curl -LsSf https://astral.sh/uv/install.sh | sh
  ```

## 環境の sync

本ツールの動作に必要な Python および依存パッケージをセットアップします。

* コンソールアプリで本ツールを展開したフォルダに移動した上で以下のコマンドを実行して下さい。

  ```shell
  uv sync
  ```

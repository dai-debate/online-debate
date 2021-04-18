# Zoom API の設定

Zoomのミーティング作成などを、プログラムから行えるようにするための手続き(API)を設定を行います。

## Zoomアカウントの作成

* 大会運営に使用する Zoom アカウントを用意して下さい。
  普段使いされている個人のアカウントでも動作に支障はありませんが、万が一ツールの誤動作で無関係のミーティングが削除されたりしないよう、[新規作成](https://zoom.us/signup)することをお勧めします。

## アプリケーションの登録

Zoom API を使用するために、アプリケーションを登録します。

* [Zoom App Marketplace](https://marketplace.zoom.us/)を開きます。
* 右上のメニューから「Sign In」をクリックし、使用するZoomアカウントでサインインして下さい。
  * 既にZoomにサインイン済みの場合にはそのまま次の手順に進んで下さい。
* 利用許諾に関するダイアログが表示されるので、内容を読んだ上で「Agree」をクリックします。
* 画面右上の「Develop」→「Build App」をクリックします。
* OAuth の所にある「Create」をクリックします。
* 「App Name」は半角英数字であれば何でも構いません。
  ここでは仮に `dkoshien` とします。
* 「Choose app type」は「Account-level app」(このアカウントが管理する複数のユーザーの操作を行う事ができる)を選択します。
* 「Would you like to publish～」は「OFF」に切り替えて下さい。
* 「Create」をクリックします。
* 作成したアプリケーションの認証情報が表示されます。
  以下の情報は後の手順で使用しますので、手元に控えてください。
  * Client ID
  * Client Secret
* Redirect URL for OAuth と Whitelist URL の両方に `https://github.com/dai-debate/online-debate` を入力してください。
* 「Continue」をクリックします。
* Basic Information の入力欄は以下の内容を入力して「Continue」をクリックします。
  * Short Description: `Online debate`
  * Long Description: `Manage meetings for online debates`
  * Company Name: 所属団体等を適宜入力してください。
    個人の場合は `Personal` とかで問題ないと思います。
    ここでは仮に `NADE` とします。
  * Name: 使用者の氏名を入力して下さい。
  * Email Address: 使用者のメールアドレスを入力して下さい。
    個人用ではなく、大会運営用に用意した Google アカウントのもので構いません。
* Add feature の入力欄になりますが、変更の必要はないので、「Continue」をクリックします。
* Add scopes の設定になりますので、以下の項目にチェックを入れた後に「Done」をクリックします。
  * Meeting → View and manage sub account's user meetings
  * User → View and manage sub account's user information
  * Room → View and manage sub account's Zoom Rooms information
* Install your app の下の「Install」をクリックします。
* 「≪App Name≫がZoomアカウントへのアクセスをリクエストしています」という画面が開くので、「認可」をクリックしてください。

[→次の手順に進む](zoom-create-user.md)

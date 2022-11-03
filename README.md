# 4d-tips-open-mail-with-editor
メールクライアント（Outlook）で新規メールを開いて編集するには

`mailto:` ([RFC2368](https://www.ietf.org/rfc/rfc2368.txt), [RFC6068](https://www.rfc-editor.org/rfc/rfc6068)) URLスキームを使えば，`OPEN URL`コマンドで手軽にメールアプリを起動することができますが，添付ファイルを追加することができません。

**注記**: [MailMate](https://manual.mailmate-app.com/extended_url_scheme)の` attachment-url`など，メールアプリが独自のパラメーターをサポートしている場合もあります。

## X-Unsent:1 

`mailto:`ではなく，MIME（.emlファイル）をメールアプリで開く，という方法があります。マルチパートのMIMEであれば，添付ファイルを含めることができます。Windowsでは，拡張ヘッダーの`X-Unsent`が設定されているメールを下書きとして編集することができますが，Macでは原則的に編集ができないようです（Apple Mail, Outlook）。

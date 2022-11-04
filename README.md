# 4d-tips-open-mail-with-editor
メールクライアント（Outlook）で新規メールを開いて編集するには

`mailto:` ([RFC2368](https://www.ietf.org/rfc/rfc2368.txt), [RFC6068](https://www.rfc-editor.org/rfc/rfc6068)) URLスキームを使えば，`OPEN URL`コマンドで手軽にメールアプリを起動することができますが，添付ファイルを追加することができません。

**注記**: [MailMate](https://manual.mailmate-app.com/extended_url_scheme)の` attachment-url`など，メールアプリが独自のパラメーターをサポートしている場合もあります。

## X-Unsent:1 

`mailto:`ではなく，MIME（.emlファイル）をメールアプリで開く，という方法があります。マルチパートのMIMEであれば，添付ファイルを含めることができます。Windowsでは，拡張ヘッダーの`X-Unsent`が設定されているメールを下書きとして編集することができますが，Macでは「差出人なし」となり，原則的に編集ができないようです（Apple Mail, Outlook）。

<img width="443" alt="" src="https://user-images.githubusercontent.com/1725068/199855384-160591db-496a-46b3-bfa9-b1a35c216c04.png">

## Outlook VBA

[Outlookのオートメーション](https://learn.microsoft.com/en-us/office/vba/api/overview/outlook)で新規メールを開いて編集することができます。

本文を[`MailItem.Body`](https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.body)にセットすると，Outlookが自動的にHTML本文を生成します。

```html
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2//EN"> 
<HTML> 
  <HEAD> 
  <META NAME="Generator" CONTENT="MS Exchange Server version 16.0.12430.20120"> 
  <TITLE></TITLE> 
  </HEAD> 
  <BODY> 
    <!-- Converted from text/rtf format --> 
    <P>
      <SPAN LANG="ja"><FONT STYLE="…">…</FONT></SPAN>
    </P> 
  </BODY> 
</HTML> 
```

スタイルは，[デフォルトのフォント](https://support.microsoft.com/en-us/office/change-or-set-the-default-font-in-outlook-20f72414-2c42-4b53-9654-d07a92b9294a)が使用されます。ただし，本文と段落にはスタイルが設定されていないため，本文に続けて入力したテキストはデフォルトのフォントではなく，Outlook既定の書式になります。

[`MailItem.HTMLBody`](https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.htmlbody)を直接セットすることもできますが，デフォルトのフォントはレジストリに保存されており，キー名はOfficeのバージョンによって違っているので，取り出すのは少し面倒です。

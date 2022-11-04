//%attributes = {}
C_OBJECT:C1216($email)

$email:=New object:C1471

$email.subject:="テストメール"
$email.textBody:="こんにちは"

$email.to:=New collection:C1472
$email.to.push(New object:C1471("name"; "わかんだちゃん"; "email"; "support@wakanda.io"))

$path:=Get 4D folder:C485(Current resources folder:K5:16)+"wakhello2.png"

$email.attachments:=New collection:C1472
$email.attachments.push(MAIL New attachment:C1644($path))

$VBA:=cs:C1710.VBA.new()

$VBA.openMailWithEditor($email)
Class extends Mail

Class constructor
	
	Super:C1705()
	
	This:C1470.scptFolder:=Folder:C1567(Get 4D folder:C485(Current resources folder:K5:16); fk platform path:K87:2).folder("scpt")
	
Function _setaddresses($addresses : Collection; $type : Text)
	
	SET ENVIRONMENT VARIABLE:C812("MAIL_"+$type+"_COUNT"; String:C10($addresses.length))
	
	For ($i; 1; $addresses.length)
		SET ENVIRONMENT VARIABLE:C812("MAIL_"+$type+"_"+String:C10($i); $addresses[$i-1])
	End for 
	
Function openMailWithEditor($email : Object)
	
	If (Is macOS:C1572)
		
		This:C1470._setExecutable("openMailInMail.scpt")
		
		If ($email#Null:C1517)
			
			$to:=This:C1470._getaddresses($email.to)
			$cc:=This:C1470._getaddresses($email.cc)
			$bcc:=This:C1470._getaddresses($email.bcc)
			
			If ($email.textBody#Null:C1517)
				
				$tempFolder:=Folder:C1567(Temporary folder:C486; fk platform path:K87:2).folder(Generate UUID:C1066)
				$tempFolder.create()
				$tempFile:=$tempFolder.file("content")
				$tempFile.setText($email.textBody; "utf-16le")
				
				SET ENVIRONMENT VARIABLE:C812("MAIL_CONTENT_PATH"; $tempFile.path)
				
			End if 
			
			If ($email.attachments#Null:C1517)
				For ($i; 1; $email.attachments.length)
					$attachment:=$email.attachments[$i-1]
					
					If ($attachment.platformPath#Null:C1517)
						$path:=$attachment.path
					Else 
						$folder:=Folder:C1567(Temporary folder:C486; fk platform path:K87:2).folder(Generate UUID:C1066)
						$folder.create()
						$file:=$folder.file($attachment.name)
						$file.setContent($attachment.getContent())
						$path:=$file.path
					End if 
					
					SET ENVIRONMENT VARIABLE:C812("MAIL_ATTACHMENT_"+String:C10($i); $path)
					
				End for 
				SET ENVIRONMENT VARIABLE:C812("MAIL_ATTACHMENT_COUNT"; String:C10($email.attachments.length))
			End if 
			
			This:C1470._setaddresses($to; "TO")
			This:C1470._setaddresses($cc; "CC")
			This:C1470._setaddresses($bcc; "BCC")
			
			SET ENVIRONMENT VARIABLE:C812("MAIL_SUBJECT"; $email.subject)
			
			This:C1470._osascript("openMailInMail.scpt"; $stdIn)
			
		End if 
	End if 
	
Function _setExecutable($fileName : Text)
	
	$file:=This:C1470.scptFolder.file($fileName)
	
	If ($file.exists)
		If (Application type:C494=4D Remote mode:K5:5)
			
			SET ENVIRONMENT VARIABLE:C812("_4D_OPTION_CURRENT_DIRECTORY"; $file.parent.platformPath)
			SET ENVIRONMENT VARIABLE:C812("_4D_OPTION_BLOCKING_EXTERNAL_PROCESS"; "true")
			LAUNCH EXTERNAL PROCESS:C811("chmod 755 "+$file.fullName)
			
		End if 
	End if 
	
Function _osascript($fileName : Text; $stdIn : Blob)->$status : Object
	
	If (Is macOS:C1572)
		
		$status:=New object:C1471
		
		$file:=This:C1470.scptFolder.file($fileName)
		
		If ($file.exists)
			
			var $stdIn; $stdOut; $stdErr : Blob
			
			SET ENVIRONMENT VARIABLE:C812("_4D_OPTION_HIDE_CONSOLE"; "true")
			
			LAUNCH EXTERNAL PROCESS:C811("osascript "+This:C1470._escape($file.path); $stdIn; $stdOut; $stdErr)
			
			$status.stdOut:=Convert to text:C1012($stdOut; "UTF-8")
			$status.stdErr:=Convert to text:C1012($stdErr; "UTF-8")
			
		End if 
	End if 
Class extends Script

Class constructor
	
	Super:C1705()
	
	This:C1470.vbsFolder:=Folder:C1567(Get 4D folder:C485(Current resources folder:K5:16); fk platform path:K87:2).folder("vbs")
	
Function _getaddresses($property : Variant)->$addresses : Collection
	
	C_COLLECTION:C1488($values)
	
	Case of 
		: (Value type:C1509($property)=Is text:K8:3)
			$values:=New collection:C1472($property)
		: (Value type:C1509($property)=Is object:K8:27)
			$values:=New collection:C1472($property)
		: (Value type:C1509($property)=Is collection:K8:32)
			$values:=$property
	End case 
	
	$addresses:=New collection:C1472
	
	If ($values#Null:C1517)
		
		C_VARIANT:C1683($value)
		
		For each ($value; $values)
			Case of 
				: (Value type:C1509($value)=Is text:K8:3)
					$addresses.push($value)
				: (Value type:C1509($value)=Is object:K8:27)
					$addresses.push($value.name+" "+"<"+$value.email+">")
			End case 
		End for each 
		
	End if 
	
Function openMailWithEditor($email : Object)
	
	If (Is Windows:C1573)
		If ($email#Null:C1517)
			
			$to:=This:C1470._getaddresses($email.to)
			$cc:=This:C1470._getaddresses($email.cc)
			$bcc:=This:C1470._getaddresses($email.bcc)
			
			If ($email.attachments#Null:C1517)
				For ($i; 1; $email.attachments.length)
					$attachment:=$email.attachments[$i-1]
					
					If ($attachment.platformPath#Null:C1517)
						$path:=$attachment.platformPath
					Else 
						$folder:=Folder:C1567(Temporary folder:C486; fk platform path:K87:2).folder(Generate UUID:C1066)
						$folder.create()
						$file:=$folder.file($attachment.name)
						$file.setContent($attachment.getContent())
						$path:=$file.platformPath
					End if 
					
					SET ENVIRONMENT VARIABLE:C812("MAIL_ATTACHMENT_"+String:C10($i); $path)
					
				End for 
			End if 
			
			SET ENVIRONMENT VARIABLE:C812("MAIL_TO"; $to.join("; "))
			SET ENVIRONMENT VARIABLE:C812("MAIL_CC"; $cc.join("; "))
			SET ENVIRONMENT VARIABLE:C812("MAIL_BCC"; $bcc.join("; "))
			SET ENVIRONMENT VARIABLE:C812("MAIL_SUBJECT"; $email.subject)
			
			C_BLOB:C604($stdIn)
			
			If ($email.htmlBody=Null:C1517)
				$template:="<!--#4dtext $1-->"
				$textBody:=""
				PROCESS 4D TAGS:C816($template; $textBody; $email.textBody)
				CONVERT FROM TEXT:C1011($textBody; "utf-16LE"; $stdIn)
				SET ENVIRONMENT VARIABLE:C812("MAIL_FORMAT"; "TEXT")
			Else 
				CONVERT FROM TEXT:C1011($email.htmlBody; "utf-16LE"; $stdIn)
				SET ENVIRONMENT VARIABLE:C812("MAIL_FORMAT"; "HTML")
			End if 
			
			This:C1470._cscript("openMailInOutlook.vbs"; $stdIn)
			
		End if 
	End if 
	
Function _cscript($fileName : Text; $stdIn : Blob)->$status : Object
	
	If (Is Windows:C1573)
		
		$status:=New object:C1471
		
		$file:=This:C1470.vbsFolder.file($fileName)
		
		If ($file.exists)
			
			var $stdIn; $stdOut; $stdErr : Blob
			SET ENVIRONMENT VARIABLE:C812("_4D_OPTION_HIDE_CONSOLE"; "true")
			
			LAUNCH EXTERNAL PROCESS:C811("cscript //Nologo //U "+This:C1470._escape_path($file.path); $stdIn; $stdOut; $stdErr)
			
			$status.stdOut:=Convert to text:C1012($stdOut; "UTF-16LE")
			$status.stdErr:=Convert to text:C1012($stdErr; "UTF-16LE")
			
		End if 
	End if 
Class constructor
	
Function openMailWithEditor($email : Object)
	
	If ($email#Null:C1517)
		
		If ($email.headers=Null:C1517)
			$email.headers:=New collection:C1472
		End if 
		
		$email.headers.push(New object:C1471("name"; "X-Unsent"; "value"; "1"))
		
		$mime:=MAIL Convert to MIME:C1604($email)
		
		$temp:=Folder:C1567(Temporary folder:C486; fk platform path:K87:2).folder(Generate UUID:C1066)
		$temp.create()
		
		If ($email.subject#Null:C1517)
			$emlFile:=$temp.file($email.subject+".eml")
		Else 
			$emlFile:=$temp.file("draft.eml")
		End if 
		
		$emlFile.setText($mime; "us-ascii"; Document with LF:K24:22)
		
		OPEN URL:C673($emlFile.platformPath)
		
	End if 
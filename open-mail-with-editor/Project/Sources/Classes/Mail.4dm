Class extends Script

Class constructor
	
	Super:C1705()
	
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
Class constructor
	
Function _escape_path($in : Text)->$out : Text
	
	If (Count parameters:C259#0)
		
		$path:=$in
		
		If (Is Windows:C1573)
			
			$path:=This:C1470._escape($path)
			
		Else 
			
			$path:=This:C1470._escape(Convert path system to POSIX:C1106($path))
			
		End if 
		
	End if 
	
	$out:=$path
	
Function _escape($in : Text)->$out : Text
	
	If (Count parameters:C259#0)
		
		$param:=$in
		
		If (Is Windows:C1573)
			
			$shoudQuote:=False:C215
			
			$metacharacters:="&|<>()%^\" "
			
			$len:=Length:C16($metacharacters)
			
			For ($i; 1; $len)
				$metacharacter:=Substring:C12($metacharacters; $i; 1)
				$shoudQuote:=$shoudQuote | (Position:C15($metacharacter; $param; *)#0)
				If ($shoudQuote)
					$i:=$len
				End if 
			End for 
			
			If ($shoudQuote)
				If (Substring:C12($param; Length:C16($param))="\\")
					$param:="\""+$param+"\\\""
				Else 
					$param:="\""+$param+"\""
				End if 
			End if 
			
		Else 
			
			$metacharacters:="\\!\"#$%&'()=~|<>?;*`[] "
			
			For ($i; 1; Length:C16($metacharacters))
				$metacharacter:=Substring:C12($metacharacters; $i; 1)
				$param:=Replace string:C233($param; $metacharacter; "\\"+$metacharacter; *)
			End for 
			
		End if 
		
	End if 
	
	$out:=$param
	
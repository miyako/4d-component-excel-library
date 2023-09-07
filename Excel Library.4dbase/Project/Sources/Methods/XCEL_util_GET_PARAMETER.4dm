//%attributes = {"shared":true}
C_POINTER:C301($1; $2)

If (Count parameters:C259=2)
	If (Not:C34(Is nil pointer:C315($1)))
		If (Not:C34(Is nil pointer:C315($2)))
			Case of 
				: (Type:C295($1->)=Is longint:K8:6) | (Type:C295($1->)=Is real:K8:4)
					Case of 
						: (Type:C295($2->)=Is longint:K8:6) | (Type:C295($2->)=Is real:K8:4)
							$1->:=$2->
						: (Type:C295($2->)=Is text:K8:3)
							$1->:=Num:C11($2->)
					End case 
				: (Type:C295($1->)=Is text:K8:3)
					Case of 
						: (Type:C295($2->)=Is longint:K8:6) | (Type:C295($2->)=Is real:K8:4)
							$1->:=String:C10($2->)
						: (Type:C295($2->)=Is text:K8:3)
							$1->:=$2->
					End case 
			End case 
		End if 
	End if 
End if 

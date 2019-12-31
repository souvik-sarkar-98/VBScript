Function CalculateSum(ByVal arr,ByVal StartingIndex, ByVal jump)
	Dim s
	s=0
	For i=StartingIndex To Len (arr) Step jump
		s=s+Mid(arr,i,1)	
	Next
	CalculateSum=s
End Function 

Function Concat(ByVal arr,ByVal StartingIndex, ByVal jump)
	Dim s
	s=""
	For i=StartingIndex To Len (arr) Step jump
		s=s&Mid(arr,i,1)	
	Next
	Concat=s
End Function

Function CheckDigit(ByVal str)
	Dim var
	If Len (str)=2 Then 
		var=Mid(str,2,1)
	Else 
	 	var=str 
	End If
	CheckDigit=var 
End Function 	
Function ValidateID(ByVal ID)
	Dim sum,var,dig
	sum=CalculateSum(ID,1,2)
	var=Concat(ID,2,2)*2
	sum=sum+CalculateSum(var,1,1)
	dig=CheckDigit(sum)
	dig=CheckDigit(10-dig)
	ValidateID=dig
End Function 

msgbox (ValidateID(Inputbox("")))
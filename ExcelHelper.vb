' <IsStraightVb>True</IsStraightVb>
Class ExcelHelper

Private GoExcel as IGoExcel
Private Parameter as IParamDynamic
Private ThisDoc As Inventor.Document
Private pathTables as String

Sub New(GoExcel As IGoExcel, Parameter As IParamDynamic, ThisDoc As Inventor.Document)
  	Me.GoExcel = GoExcel  
	Me.Parameter = Parameter
	Me.ThisDoc = ThisDoc
	Me.pathTables = Environ("CAD_SPREADSHEETS")
	If Not (pathTables.EndsWith("/") Or pathTables.EndsWith("\\")) Then
		pathTables = pathTables & "/"
	End If
End Sub

Function TranslateCode(codeTable As String, sheet As String, code As String, codeCol as String, valueCol as String) As String    
Dim codeReturn As String    
Dim table = pathTables & codeTable & ".xlsx"    
GoExcel.Open(table, sheet) 
GoExcel.TitleRow = 1

Dim i As Integer = GoExcel.FindRow(table, sheet, codeCol, "=", code)    
If i = -1 Then        
	codeReturn = code    
Else   
      codeReturn = GoExcel.CurrentRowValue(valueCol)    
End If    
GoExcel.Close        
Return codeReturn
End Function


Function SetParameters(codeTable As String,  sheet As String, code As String, codeCol as String) As Boolean    

Dim table = pathTables & codeTable & ".xlsx"    
GoExcel.Open(table, sheet) 
GoExcel.TitleRow = 1

Dim i As Integer = GoExcel.FindRow(table, sheet, codeCol, "=", code)    
If i = -1 Then  
	GoExcel.Close 
	Return False  
Else 
    Dim oParams As Inventor.Parameters
    oParams = ThisDoc.ComponentDefinition.Parameters

    ' Iterate through the Parameters collection to obtain
    ' information about the Parameters
    Dim iNumParams As Long
    For iNumParams = 1 To oParams.Count
        If oParams.Item(iNumParams).ParameterType = Inventor.ParameterTypeEnum.kUserParameter Then
		
        	Dim name As String = oParams.Item(iNumParams).Name
			
			Dim newValue As String = GoExcel.CurrentRowValue(name)
	
        If Not (newValue Is Nothing Or newValue.Trim().Length = 0) Then
				Log("Units=" & oParams.Item(iNumParams).Units)
			Select Case oParams.Item(iNumParams).Units
			Case "Boolean" 
				If  (newValue.StartsWith("Y") Or _
			  									newValue.StartsWith("y") Or _
												newValue.ToLower() = "true" Or _
												newValue = "1") Then
					Parameter(name) = True
				Else
					Parameter(name) = False
				End If
				Exit Select
			Case "Text"
				Parameter(name) = newValue
				Exit Select
			Case Else
				Parameter(name) = Val(newValue)
				
			End Select
    		End If
		End If
   Next
   
End If    
GoExcel.Close        
Return True
End Function

Function GetPath () As String
	Return pathTables
End Function

Private theLog as System.IO.StreamWriter
Private Sub CreateLog()
        Dim fileNameExt As String = "c:/temp/ExcelHelper.log"
        Dim fs As New System.IO.FileStream(fileNameExt, System.IO.FileMode.Append, System.IO.FileAccess.Write)
        theLog =  New System.IO.StreamWriter(fs)
    End Sub
  Private  Sub Log(ByVal msg As String)
		If theLog Is Nothing Then
			CreateLog()
		End If
 
            theLog.WriteLine(DateAndTime.Now().ToString() & "   " & msg)
            theLog.Flush()
     End Sub
End Class

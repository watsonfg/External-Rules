' <IsStraightVb>True</IsStraightVb>
' <FireOthersImmediately>False</FireOthersImmediately>
Imports System.Diagnostics
Class WatsonFunction

    Public Function GetCuttopErpByEdge(ByVal PassedEDGTYP As String) As String
        Try
            If PassedEDGTYP = "" Then
                ' return a value with ERROR prefix.
                Return "ERROR- No EDGTYP was passed. (I: WatsonFunction GetCuttopErpByEdge)\n"
            End If
        Catch ex As Exception
            Log("EXCEPTION thrown in GetCuttopErpByEdge(" & PassedEDGTYP & "): " & ex.Message)
            Log("STACK:" & ex.StackTrace)
            Return Nothing
        End Try

        Try
            'Create the Return value here...
            Select Case PassedEDGTYP
                Case "AR", "AW", "AY", "BF", "BH", "HG", "DB", "DC", "DD", "DR", "DT", "WE", "WR", "WS"
                    Return 710
                Case "PE", "PR", "PS", "F", "M", "N", "S", "T", "U", "K", "V", "W"
                    Return 411
                Case "B", "X"
                    Return 700
                Case "Q", "Z", "C", "J", "HG"
                    Return 510
                Case Else
					' return a value with ERROR prefix.
                    Return "ERROR- Passed EDGTYP of: '" & PassedEDGTYP & "' is NOT a Valid Edge Type. (I: WatsonFunction GetCuttopErpByEdge)\n"
            End Select
        Catch ex As Exception
            Return ("EXCEPTION thrown in GetCuttopErpByEdge(" & PassedEDGTYP & "): " & ex.Message)
        End Try
    End Function
	
	Public Function GetMattypByEdge(ByVal EDGTYP As String, ByVal Reversible As Boolean) As String
        Try
            If EDGTYP = "" Then
                ' return a value with ERROR prefix.
                Return "ERROR - No EDGTYP was passed. (I: WatsonFunction GetMattypByEdge)\n"
            End If
        Catch ex As Exception
            Log("EXCEPTION thrown in GetMattypByEdge(" & EDGTYP & "): " & ex.Message)
            Log("STACK:" & ex.StackTrace)
            Return Nothing
        End Try

        Try
			If Reversible Then
				'Create the Return value here...
				Select Case EDGTYP
					Case "K", "U", "J", "S", "N"
						Return "MP1"
					Case "Q", "Z", "X"
						Return "P1"
					Case Else
						' return a value with ERROR prefix.
						Return "ERROR - EDGTYP of: '" & EDGTYP & "' is NOT a Valid Edge Type. (I: WatsonFunction GetMattypByEdge)\n"
				End Select
			Else
				'Create the Return value here...
				Select Case EDGTYP
					Case "K", "U", "J", "S", "N"
						Return "MP1"
					Case "Q", "Z", "X"
						Return "P1"
					Case Else
						' return a value with ERROR prefix.
						Return "ERROR - EDGTYP of: '" & EDGTYP & "' is NOT a Valid Edge Type. (I: WatsonFunction GetMattypByEdge)\n"
				End Select
			End If		
        Catch ex As Exception
            Return ("EXCEPTION thrown in GetMattypByEdge(" & PassedEDGTYP & "): " & ex.Message)
        End Try
    End Function
	
''' Below is all of the DebugHelper code Commented out as reference, GH 4/9/15

'    Private ThisDoc As Inventor.Document
'    Private debugOn As Boolean
'
'    Private Shared fileNameExt As String
'    Sub New(ThisDoc As Inventor.Document, Optional SharedVariable As ISharedVariable = Nothing)
'        Dim process As Process = Process.GetCurrentProcess()
'
'        Me.ThisDoc = ThisDoc
'        Dim sDebug As String = Environ("ILOGIC_DEBUG")
'
'        If Me.fileNameExt Is Nothing Then
'            If Not SharedVariable Is Nothing Then
'                If SharedVariable.Exists("A__SERVER_LOG") Then
'                    Me.fileNameExt = SharedVariable("A__SERVER_LOG")
'                End If
'            End If
'
'            If fileNameExt Is Nothing Then
'                Me.fileNameExt = GetProperty("A__SERVER_LOG")
'                If Not SharedVariable Is Nothing And Not Me.fileNameExt Is Nothing Then
'                    SharedVariable("A__SERVER_LOG") = Me.fileNameExt
'                End If
'            End If
'
'        End If
'        If fileNameExt Is Nothing Then
'            Dim logFilePath As String = Environ("ILOGIC_LOG")
'
'            If logFilePath Is Nothing Then
'                logFilePath = "c:\temp\"
'            End If
'            If Not (logFilePath.EndsWith("/") Or logFilePath.EndsWith("\\")) Then
'                logFilePath = logFilePath & "/"
'            End If
'            If Not SharedVariable Is Nothing Then
'                Me.fileNameExt = logFilePath & "iLogic.log"
'            Else
'                Me.fileNameExt = logFilePath & "iLogic_" & process.Id & ".log"
'            End If
'
'        End If
'
'        Me.debugOn = Not sDebug Is Nothing And sDebug = "true"
'        'Log("ENTERING ..." & ThisDoc.DisplayName)'GH replaced simple file name with full path and file
'        Log("ENTERING ..." & ThisDoc.FullFileName)
'    End Sub
'
'    '''GH turned this off on 2/14/13 since we are not really looking at the data in the logs.
    'Sub DumpParameters()     
    '		If Not Me.DebugOn Then
    '			Exit Sub
    '		End If
    '   Dim fs As New System.IO.FileStream(fileNameExt, System.IO.FileMode.Append, System.IO.FileAccess.Write)
    '     Dim theLog As New System.IO.StreamWriter(fs)
    '
    '    ' Iterate through the Parameters collection to obtain
    '    ' information about the Parameters
    '	theLog.WriteLine(DateAndTime.Now().ToString() & " Parameters for " & ThisDoc.DisplayName & ":")
    '	
    '	Dim oParams As Inventor.Parameters
    '    oParams = ThisDoc.ComponentDefinition.Parameters
    '
    '    Dim iNumParams As Long
    '    For iNumParams = 1 To oParams.Count
    '         	Dim name As String = oParams.Item(iNumParams).Name
    '			Dim value As String = oParams.Item(iNumParams).Value
    '			theLog.WriteLine(DateAndTime.Now().ToString() & " .... " & name & "=" & value)
    '		   Next
    '	theLog.Close()
    ' End Sub

'    Sub DumpProperties()
'        If Not Me.DebugOn Then
'            Exit Sub
'        End If
'        Dim fs As New System.IO.FileStream(fileNameExt, System.IO.FileMode.Append, System.IO.FileAccess.Write)
'        Dim theLog As New System.IO.StreamWriter(fs)
'
'        ' Iterate through the Parameters collection to obtain
'        ' information about the Parameters
'        theLog.WriteLine(DateAndTime.Now().ToString() & " Custom iProperties for " & ThisDoc.DisplayName & ":")
'        Dim oPropSet As Inventor.PropertySet
'        oPropSet = ThisDoc.PropertySets.Item("Inventor User Defined Properties")
'        Dim p As Long
'        For p = 1 To oPropSet.Count
'            Dim oCustomProp As Inventor.Property = oPropSet(p)
'            Dim name As String = oCustomProp.Name
'            Dim value As String = oCustomProp.Value
'
'            theLog.WriteLine(DateAndTime.Now().ToString() & " .... " & name & "=" & value)
'        Next
'        theLog.Close()
'    End Sub

'    Sub Log(ByVal msg As String)
'        If Not Me.DebugOn Then
'            Exit Sub
'        End If
'        Try
'            Dim fs As New System.IO.FileStream(fileNameExt, System.IO.FileMode.Append, System.IO.FileAccess.Write)
'            Dim theLog As New System.IO.StreamWriter(fs)
'            theLog.WriteLine(DateAndTime.Now().ToString() & "   " & msg)
'            theLog.Close()
'        Catch ex As Exception
'        End Try
'
'    End Sub

'    Private Function GetProperty(ByVal name As String) As String
'        ' Access a particular property set.  In this case the design tracking property set.
'        Dim oDTProps As Inventor.PropertySet
'
'        ' get the Custom Properties
'        Try
'            oDTProps = ThisDoc.PropertySets.Item("Inventor User Defined Properties")
'        Catch ex As Exception
'            Log("EXCEPTION thrown in GetProperty(" & name & "): " & ex.Message)
'            Log("STACK:" & ex.StackTrace)
'            Return Nothing
'        End Try
'
'
'        ' Get a specific property, in this case the designer property.
'        Dim oCustomProp As Inventor.Property = Nothing
'
'        ' You can also use the name or display name, the display name has the problem that it can be changed.
'        Try
'            oCustomProp = oDTProps.Item(name)
'        Catch ex As Exception
'            ' Log(0, "EXCEPTION thrown in GetProperty(" & name & "): " & ex.Message)
'            Return Nothing
'        End Try
'        ' return the value.
'        Return oCustomProp.Value
'
'    End Function

End Class


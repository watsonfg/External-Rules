' <IsStraightVb>True</IsStraightVb>
Imports System.Diagnostics
'
'  	Initialize by creating new application and calling with ThisApplication
'
'  	Properties
'		Value - (Double) 			Returns current RLOAD value Of Part
'		CountOfHoles - (Integer) 	Returns the Number Of Hole Features in current part
'		IsInErrorState - (Boolean)	Returns True if an Error has been triggered
'		ErrorValue - (String)		Returns error message(s) that generated the error state
'		
'	SubRoutines
'		 (Not Ready)FeedRate()					Displays MessageBox With Feedrate And material type Of current part
'		 Calculate()				Determine the current RLOAD value And populates the RLOAD segment
'		 ShowRLOADFace()			Displays calculation Of RLOAD value And highlight the correct face that Is being used

Public Class RLOADasm
    Private _dRLOAD As Double = -999.0
    Private _oDoc As Inventor.AssemblyDocument 'was PartDocument
    Private _oApp As Inventor.Application
    Private _constCMtoInch As Double = 0.393700787
    Private _dPerimeter As Double = 0
	Private _errorState As Boolean = False
	Private _strErrorValue As String = ""
	Private _iProperties As Inventor.PropertySet
	Private _selectedFaceNum As Integer = 0
    Private _partName As String = ""
    Private _corePartName As String = "CorePart"
    Private _bIsFaceGroup As Boolean = False
    Private _iFaceGroupQty As Integer = 0
	
	Public iWARNING As Inventor.Property
	Public iERROR As Inventor.Property
	
    Sub New(ByVal ThisDoc As Inventor.Document)
       If ThisDoc.DocumentType = Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then 'was kPartDocumentObject
            _oApp = ThisApplication
			_oDoc = ThisDoc

            On Error Resume Next
			iProperties = _oDoc.PropertySets.Item("Inventor User Defined Properties")

            iERROR = iProperties.Item("ERRORS")
            If Err.Number <> 0 Then
                _errorState = True
                _strErrorValue = _strErrorValue & "\\nNo ERRORS iProperty Defined for this Part"
				
				Err.Clear
            End If			
			

            iWARNING = iProperties.Item("WARNINGS")
            If Err.Number <> 0 Then
                _errorState = True
                _strErrorValue = _strErrorValue & "\\nNo WARNINGS iProperty Defined for this Part"
				Err.Clear
            End If

            Dim invDesignInfo As Inventor.PropertySet
            invDesignInfo = _oDoc.PropertySets.Item("Design Tracking Properties")
            _partName = invDesignInfo.Item("Part Number").Value
        Else
            _errorState = True
            _strErrorValue = "This rule is not being run in an IPT file."
        End If
    End Sub

	
	
	Public Sub CheckErr()
        Dim ipsetCustom As Inventor.PropertySet
		ipsetCustom = _oDoc.PropertySets.Item("Inventor User Defined Properties")
        On Error Resume Next
	   iERROR = ipsetCustom.Item("ERRORS")
        If Err.Number <> 0 Then
            _errorState = True
            _strErrorValue = _strErrorValue & "\\nNo ERRORS iProperty Defined for this Part"
				AddErr()
            Err.Clear()
        End If
    End Sub
	
		
	Public Sub CheckWarn()
        Dim ipsetCustom As Inventor.PropertySet
		ipsetCustom = _oDoc.PropertySets.Item("Inventor User Defined Properties")
        On Error Resume Next
		iWARNINGS = ipsetCustom.Item("WARNINGS")
        If Err.Number <> 0 Then
            _errorState = True
            _strErrorValue = _strErrorValue & "\\nNo WARNINGS iProperty Defined for this Part"
				AddWarn()
            Err.Clear()
        End If
    End Sub
	
	
	Public Sub CheckRload()
        Dim ipsetCustom As Inventor.PropertySet
		ipsetCustom = _oDoc.PropertySets.Item("Inventor User Defined Properties")
        On Error Resume Next
		iRLOAD = ipsetCustom.Item("RLOAD")
        If Err.Number <> 0 Then
            _errorState = True
            _strErrorValue = _strErrorValue & "\\nNo RLOAD iProperty Defined for this Part"
				AddRload()
            Err.Clear()
        End If
    End Sub
	
		Public Sub CheckLload()
        Dim ipsetCustom As Inventor.PropertySet
		ipsetCustom = _oDoc.PropertySets.Item("Inventor User Defined Properties")
        On Error Resume Next
		iLLOAD = ipsetCustom.Item("LLOAD")
        If Err.Number <> 0 Then
            _errorState = True
            _strErrorValue = _strErrorValue & "\\nNo LLOAD iProperty Defined for this Part"
				AddLload()
            Err.Clear()
        End If
    End Sub
	
	
	Public Sub CheckSqFt()
        Dim ipsetCustom As Inventor.PropertySet
		ipsetCustom = _oDoc.PropertySets.Item("Inventor User Defined Properties")
        On Error Resume Next
		iRLOAD = ipsetCustom.Item("SQFT")
        If Err.Number <> 0 Then
            _errorState = True
            _strErrorValue = _strErrorValue & "\\nNo SQFT iProperty Defined for this Part"
				AddSqft()
            Err.Clear()
        End If
    End Sub
	
	
	Public Sub CheckEdgSeq()
        Dim ipsetCustom As Inventor.PropertySet
		ipsetCustom = _oDoc.PropertySets.Item("Inventor User Defined Properties")
        On Error Resume Next
		iEDGSEQ = ipsetCustom.Item("EDGSEQ")
        If Err.Number <> 0 Then
            _errorState = True
            _strErrorValue = _strErrorValue & "\\nNo EDGSEQ iProperty Defined for this Part"
				AddEdgSeq()
            Err.Clear()
        End If
    End Sub
	
	
	Public Sub CheckEdgLft()	'Conditional on Edge Sequence
			Dim ipsetCustom As Inventor.PropertySet
			ipsetCustom = _oDoc.PropertySets.Item("Inventor User Defined Properties")		
				iEDGSEQ = ipsetCustom.Item("EDGSEQ")
					If iEDGSEQ.Value <> "NO EDGE" Then
			
						On Error Resume Next
						iEDGLFT = ipsetCustom.Item("EDGLFT")
						If Err.Number <> 0 Then
							_errorState = True
							_strErrorValue = _strErrorValue & "\\nNo EDGLFT iProperty Defined for this Part"
								AddEdgLft()
							Err.Clear()
						End If		
					End If
    End Sub		
	
	
	
    Public Sub AddErr()
        Dim ipsetCustom As Inventor.PropertySet
		ipsetCustom = _oDoc.PropertySets.Item("Inventor User Defined Properties")
        Dim iholder As Inventor.Property
        Dim strText As String = ""
        iholder = ipsetCustom.Add(strText, "ERRORS")
		_oDoc.Save 'force update
		_errorState = False 'reset
		'recheck
		On Error Resume Next
	   iERROR = ipsetCustom.Item("ERRORS")
        If Err.Number <> 0 Then
            _errorState = True
            _strErrorValue = _strErrorValue & "\\nNo ERRORS iProperty Defined for this Part"
            Err.Clear()
        End If
		
    End Sub
	
	
    Public Sub AddWarn()
        Dim ipsetCustom As Inventor.PropertySet
		ipsetCustom = _oDoc.PropertySets.Item("Inventor User Defined Properties")
        Dim iholder As Inventor.Property
        Dim strText As String = ""
        iholder = ipsetCustom.Add(strText, "WARNINGS")
		_oDoc.Save 'force update
		_errorState = False 'reset
		'recheck
        On Error Resume Next
		iWARNINGS = ipsetCustom.Item("WARNINGS")
        If Err.Number <> 0 Then
            _errorState = True
            _strErrorValue = _strErrorValue & "\\nNo WARNINGS iProperty Defined for this Part"
            Err.Clear()
        End If
    End Sub
	
	
    Public Sub AddRload()
        Dim ipsetCustom As Inventor.PropertySet
		ipsetCustom = _oDoc.PropertySets.Item("Inventor User Defined Properties")
        Dim iholder As Inventor.Property
        Dim dNum As Double = 0
        iholder = ipsetCustom.Add(dNum, "RLOAD")
		_oDoc.Save 'force update
		_errorState = False 'reset
        On Error Resume Next
		iRLOAD = ipsetCustom.Item("RLOAD")
        If Err.Number <> 0 Then
            _errorState = True
            _strErrorValue = _strErrorValue & "\\nNo RLOAD iProperty Defined for this Part"
            Err.Clear()
        End If
    End Sub
	
	
	Public Sub AddLload()
        Dim ipsetCustom As Inventor.PropertySet
		ipsetCustom = _oDoc.PropertySets.Item("Inventor User Defined Properties")
        Dim iholder As Inventor.Property
        Dim dNum As Double = 0
        iholder = ipsetCustom.Add(dNum, "LLOAD")
		_oDoc.Save 'force update
		_errorState = False 'reset
        On Error Resume Next
		iLLOAD = ipsetCustom.Item("LLOAD")
        If Err.Number <> 0 Then
            _errorState = True
            _strErrorValue = _strErrorValue & "\\nNo LLOAD iProperty Defined for this Part"
            Err.Clear()
        End If
    End Sub
	
		
    Public Sub AddSqFt()
        Dim ipsetCustom As Inventor.PropertySet
		ipsetCustom = _oDoc.PropertySets.Item("Inventor User Defined Properties")
        Dim iholder As Inventor.Property
		Dim dNum As Double = 0
        iholder = ipsetCustom.Add(dNum, "SQFT")
		_oDoc.Save 'force update
		_errorState = False 'reset
		'recheck	
        On Error Resume Next
		iSQFT = ipsetCustom.Item("SQFT")
        If Err.Number <> 0 Then
            _errorState = True
            _strErrorValue = _strErrorValue & "\\nNo SQFT iProperty Defined for this Part"
            Err.Clear()
        End If
    End Sub	
	
	
    Public Sub AddEdgSeq()
        Dim ipsetCustom As Inventor.PropertySet
		ipsetCustom = _oDoc.PropertySets.Item("Inventor User Defined Properties")
        Dim iholder As Inventor.Property
		Dim strText As String = ""
        iholder = ipsetCustom.Add(strText, "EDGSEQ")
		_oDoc.Save 'force update
		_errorState = False 'reset
		'recheck	
        On Error Resume Next
		iEDGSEQ = ipsetCustom.Item("EDGSEQ")
        If Err.Number <> 0 Then
            _errorState = True
            _strErrorValue = _strErrorValue & "\\nNo EDGSEQ iProperty Defined for this Part"
            Err.Clear()
        End If
    End Sub	
	
	
    Public Sub AddEdgLft()
        Dim ipsetCustom As Inventor.PropertySet
		ipsetCustom = _oDoc.PropertySets.Item("Inventor User Defined Properties")
        Dim iholder As Inventor.Property
		Dim dNum As Double = 0
        iholder = ipsetCustom.Add(dNum, "EDGLFT")
		_oDoc.Save 'force update
		_errorState = False 'reset
		'recheck	
        On Error Resume Next
		iEDGLFT = ipsetCustom.Item("EDGLFT")
        If Err.Number <> 0 Then
            _errorState = True
            _strErrorValue = _strErrorValue & "\\nNo EDGLFT iProperty Defined for this Part"
            Err.Clear()
        End If
    End Sub	
	
	
End Class

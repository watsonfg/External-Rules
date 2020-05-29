' <IsStraightVb>True</IsStraightVb>
Imports System.Diagnostics
'
'  	Initialize by creating new application and calling with ThisApplication
'
'  	Properties
'		Value - (Double) 			        Returns current LLOAD value Of Part
'		CountOfHoles - (Integer) 	        Returns the Number Of Hole Features in current part
'		IsInErrorState - (Boolean)	        Returns True if an Error has been triggered
'		ErrorValue - (String)		        Returns error message(s) that generated the error state
'		
'	SubRoutines
'		 (Not Ready)FeedRate()				Displays MessageBox With Feedrate And material type Of current part
'		 Calculate()				        Determine the current LLOAD value And populates the LLOAD segment
'		 (Not Ready)ShowLLOADFace()			Displays calculation Of LLOAD value And highlight the correct face that Is being used

Public Class LLOAD
    Private _dLLOAD As Double = -999.0
    Private _oDoc As Inventor.PartDocument
    Private _oApp As Inventor.Application
    Private _constCMtoInch As Double = 0.393700787
    Private _dPerimeter As Double = 0
    Private _errorState As Boolean = False
    Private _strErrorValue As String = ""
    Private _iProperties As Inventor.PropertySet

    Public iWARNING As Inventor.Property
    Public iERROR As Inventor.Property
    Public iLLOAD As Inventor.Property

    Sub New(ByVal ThisDoc As Inventor.Document)
        If ThisDoc.DocumentType = Inventor.DocumentTypeEnum.kPartDocumentObject Then
            _oApp = ThisApplication
            _oDoc = ThisDoc

            On Error Resume Next
            iProperties = _oDoc.PropertySets.Item("Inventor User Defined Properties")

            iERROR = iProperties.Item("ERRORS")
            If Err.Number <> 0 Then
                _errorState = True
                _strErrorValue = _strErrorValue & "\\nNo ERRORS iProperty Defined for this Part"
                Err.Clear()
            End If


            iWARNING = iProperties.Item("WARNINGS")
            If Err.Number <> 0 Then
                _errorState = True
                _strErrorValue = _strErrorValue & "\\nNo WARNINGS iProperty Defined for this Part"
                Err.Clear()
            End If

            iLLOAD = iProperties.Item("LLOAD")
            If Err.Number <> 0 Then
                _errorState = True
                _strErrorValue = _strErrorValue & "\\nNo LLOAD iProperty Defined for this Part"
                Err.Clear()
            End If

            Dim invDesignInfo As Inventor.PropertySet
            invDesignInfo = _oDoc.PropertySets.Item("Design Tracking Properties")
            _partName = invDesignInfo.Item("Part Number").Value
        Else
            _errorState = True
            _strErrorValue = "This rule is not being run in an IPT file."
        End If
    End Sub

    ReadOnly Property Value() As Double
        Get
            Return _dLLOAD
        End Get
    End Property

    Property IsInErrorState() As Boolean
        Get
            Return _errorState
        End Get
        Set(ByVal value As Boolean)
            If Not value Then
                _errorState = value
                _strErrorValue = ""
            End If
        End Set
    End Property

    ReadOnly Property ErrorValue() As String
        Get
            Return _strErrorValue
        End Get
    End Property

    Private Function _getPerimeter()
        Dim dOutside As Double = 0
        Dim dOutsideT As Double = 0
        Dim dOutWrk As Double = 0
        Dim dInside As Double = 0
        Dim dInsideT As Double = 0
        Dim dInWrk As Double = 0
        Dim iLoop As Integer
        Dim dTempPerim As Double = 0
        Dim sheetMetalDef As Inventor.SheetMetalComponentDefinition = _oDoc.ComponentDefinition

        If (Not sheetMetalDef.HasFlatPattern()) Then
            sheetMetalDef.Unfold()                          ' Warning: This may not choose the right face to unfold.  It is better if the flat pattern has been created already.
        End If

        For iLoop = 1 To sheetMetalDef.FlatPattern.Body.Faces.Count
            _ShowFacePerimeter(sheetMetalDef.FlatPattern.Body.Faces(iLoop), dOutWrk, dInWrk)
            If dOutWrk >= dOutside Then
                dOutside = dOutWrk
                If (dOutWrk + dInWrk) > dTempPerim Then
                    dTempPerim = dOutWrk + dInWrk
                    _selectedFaceNum = iLoop
                End If
            End If
            If dInWrk >= dInside Then
                dInside = dInWrk
            End If

            dOutWrk = 0
            dInWrk = 0
        Next iLoop

        _getPerimeter = (dOutside + dInside)
    End Function

    Private Function _FeedRate()
        Dim iProperties As Inventor.PropertySet
        iProperties = _oDoc.PropertySets.Item("Inventor User Defined Properties")
        Dim iMATTYP As Inventor.Property

        On Error Resume Next
        iMATTYP = iProperties.Item("MATTYP")
        If Err.Number <> 0 Then
            iERROR.Value = iERROR.Value & "\\nNo MATTYP iProperty defined [I: " & _partName & "]"
            _errorState = True
        Else
            Select Case iMATTYP.Value		'MATTYP's in descending order for best results. Values are Inches per Minute
                Case "500A36PO"
                    _FeedRate = 60 / 5		'Was 30/5 - Revised 02.20.2019
                Case "250A36PO"
                    _FeedRate = 90 / 5		'Was 70/5 - Revised 02.20.2019
                Case "22GACRS"
                    _FeedRate = 525 / 5		'Was 400/5 - Revised 02.20.2019
                Case "20GACRS"
                    _FeedRate = 370 / 5
				Case "20PERF"
                    _FeedRate = 370 / 5
				Case "20PERFGAL"			'20GA GALVANNEAL PERF 
                    _FeedRate = 370 / 5
                Case "18GACRS"
                    _FeedRate = 350 / 5		'Was 280/5 - Revised 02.20.2019
                Case "18PERF"
                    _FeedRate = 350 / 5		'Added 07.25.2019
                Case "188A36PO"
                    _FeedRate = 180 / 5		'Was 90/5 - Revised 02.20.2019
                Case "16GAAL"
                    _FeedRate = 260 / 5 	'16GA Aluminum
                Case "16GACRS"
                    _FeedRate = 220 / 5		'Was 240/5 - Revised 02.20.2019
                Case "14GACRS"
                    _FeedRate = 180 / 5
                Case "12GACRS"
                    _FeedRate = 450 / 5		'Was 150/5 - Revised 02.20.2019
                Case "11GACRS"
                    _FeedRate = 120 / 5
                Case "10GACRS"
                    _FeedRate = 80 / 5		'Was 100/5 - Revised 02.20.2019
                Case "8GAAL"
                    _FeedRate = 100 / 5	
				
                Case Else
                    iWARNING.Value = iWARNING.Value & "\\nNo FeedRate defined for " & iMATTYP.Value & " [I: " & _partName & "]"
                    _errorState = True
                    _FeedRate = 1
            End Select
        End If
    End Function

    Public Sub Calculate()
        Dim iProperties As Inventor.PropertySet
        iProperties = _oDoc.PropertySets.Item("Inventor User Defined Properties")
        Dim iLLOAD As Inventor.Property

        On Error Resume Next
        iLLOAD = iProperties.Item("LLOAD")
        If Err.Number <> 0 Then
            _errorState = True
            iERROR.Value = iERROR.Value & "\\nNo LLOAD defined [I: " & _partName & "]"
        Else
            _dPerimeter = _getPerimeter() * _constCMtoInch    ' Convert to Inches from cm
            _dLLOAD = _dPerimeter / _FeedRate()

            iLLOAD.Value = Round(_dLLOAD, 4)
        End If
    End Sub

    Private Sub _ShowFacePerimeter(ByVal oFace As Inventor.Face, ByRef dOutside As Double, ByRef dInside As Double)
        ' Find the outer loop.
        Dim dOuterLength As Double = 0
        Dim oLoop As Inventor.EdgeLoop
        Dim dMin As Double
        Dim dMax As Double
        Dim dLength As Double
        For Each oLoop In oFace.EdgeLoops
            If oLoop.IsOuterEdgeLoop Then
                Dim oEdge As Inventor.edge
                For Each oEdge In oLoop.Edges
                    Call oEdge.Evaluator.GetParamExtents(dMin, dMax)
                    Call oEdge.Evaluator.GetLengthAtParam(dMin, dMax, dLength)
                    dOuterLength = dOuterLength + dLength
                Next
                dOutside = dOuterLength
                Exit For
            End If
        Next
        ' Iterate through the inner loops.
        Dim iLoopCount As Long
        iLoopCount = 0
        Dim dTotalLength As Double
        For Each oLoop In oFace.EdgeLoops
            Dim dLoopLength As Double
            dLoopLength = 0
            If Not oLoop.IsOuterEdgeLoop Then
                For Each oEdge In oLoop.Edges
                    ' Get the length of the current edge.
                    Call oEdge.Evaluator.GetParamExtents(dMin, dMax)
                    Call oEdge.Evaluator.GetLengthAtParam(dMin, dMax, dLength)
                    dLoopLength = dLoopLength + dLength
                Next
                ' Add this loop to the total length.
                dTotalLength = dTotalLength + dLoopLength
            End If
        Next
        dInside = dTotalLength
    End Sub

    'Public Sub ShowLLOADFace()
    '    Call _getPerimeter()
    '    Dim oFace As Inventor.Face
    '    oFace = _oDoc.ComponentDefinition.Features.ExtrudeFeatures("CorePart").Faces(_selectedFaceNum)
    '    ' Clear the select set so it doesn't interfere with the highlight.
    '    _oDoc.SelectSet.Clear()

    '    ' Set a reference to the UnitsOfMeasure object to use
    '    ' in converting the values obtained to the current
    '    ' document units.  The lengths returned by the API will
    '    ' always be in centimeters.
    '    Dim oUOM As Inventor.UnitsOfMeasure
    '    oUOM = _oDoc.UnitsOfMeasure

    '    ' Create a string that will contain the loop information.
    '    Dim strResults As String

    '    ' Create the highlight sets.
    '    Dim oOuterHS As Inventor.HighlightSet
    '    oOuterHS = _oDoc.CreateHighlightSet
    '    oOuterHS.Color = _oApp.TransientObjects.CreateColor(255, 0, 0)
    '    Dim oInnerHS As Inventor.HighlightSet
    '    oInnerHS = _oDoc.CreateHighlightSet
    '    oInnerHS.Color = _oApp.TransientObjects.CreateColor(255, 255, 0)
    '    Dim dMin As Double, dMax As Double
    '    Dim dLength As Double

    '    ' Find the outer loop.
    '    Dim dOuterLength As Double
    '    dOuterLength = 0
    '    Dim oLoop As Inventor.EdgeLoop
    '    For Each oLoop In oFace.EdgeLoops
    '        If oLoop.IsOuterEdgeLoop Then
    '            Dim oEdge As Inventor.Edge
    '            For Each oEdge In oLoop.Edges
    '                ' Add this edge to the outer highlight set.
    '                oOuterHS.AddItem(oEdge)
    '                ' Get the length of the current edge.
    '                Call oEdge.Evaluator.GetParamExtents(dMin, dMax)
    '                Call oEdge.Evaluator.GetLengthAtParam(dMin, dMax, dLength)
    '                dOuterLength = dOuterLength + dLength
    '            Next
    '            ' Add the to the result message string.
    '            strResults = "Outer Loop Length (red): " & oUOM.GetStringFromValue(dOuterLength, Inventor.UnitsTypeEnum.kDefaultDisplayLengthUnits) & vbCrLf & vbCrLf
    '            Exit For
    '        End If
    '    Next

    '    ' Iterate through the inner loops.
    '    Dim iLoopCount As Long
    '    iLoopCount = 0
    '    Dim dTotalLength As Double
    '    For Each oLoop In oFace.EdgeLoops
    '        Dim dLoopLength As Double
    '        dLoopLength = 0
    '        If Not oLoop.IsOuterEdgeLoop Then
    '            For Each oEdge In oLoop.Edges
    '                ' Add this edge to the inner highlight set.
    '                oInnerHS.AddItem(oEdge)
    '                ' Get the length of the current edge.
    '                Call oEdge.Evaluator.GetParamExtents(dMin, dMax)
    '                Call oEdge.Evaluator.GetLengthAtParam(dMin, dMax, dLength)
    '                dLoopLength = dLoopLength + dLength
    '            Next
    '            ' Add this loop to the total length.
    '            dTotalLength = dTotalLength + dLoopLength
    '            ' Add to the result message string.
    '            iLoopCount = iLoopCount + 1
    '            strResults = strResults & "Inner Loop " & iLoopCount & " Length: " & _
    '              oUOM.GetStringFromValue(dLoopLength, Inventor.UnitsTypeEnum.kDefaultDisplayLengthUnits) & vbCrLf
    '        End If
    '    Next
    '    ' Display the results.
    '    strResults = strResults & "Total Inner Loop Length (yellow): " & oUOM.GetStringFromValue(dTotalLength, Inventor.UnitsTypeEnum.kDefaultDisplayLengthUnits)
    '    strResults = strResults & vbCrLf & vbCrLf & "Total Inner and Outer Loop Length: " & oUOM.GetStringFromValue(dTotalLength + dOuterLength, Inventor.UnitsTypeEnum.kDefaultDisplayLengthUnits)
    '    MsgBox(strResults)
    'End Sub

    'Public Sub FeedRate()
    '    Dim strResults As String = ""

    '    strResults = "MATTYP:  " & _iProperties.Item("MATTYP").Value & vbCrLf & vbCrLf
    '    strResults = strResults & "FeedRate:  " & _feedRate()

    '    MsgBox(strResults)
    'End Sub




'GH added subs to check and add iProperties
	
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
	
		
'	Public Sub CheckSqft()
'        Dim ipsetCustom As Inventor.PropertySet
'		ipsetCustom = _oDoc.PropertySets.Item("Inventor User Defined Properties")
'        On Error Resume Next
'		iSQFT = ipsetCustom.Item("SQFT")
'        If Err.Number <> 0 Then
'            _errorState = True
'            _strErrorValue = _strErrorValue & "\\nNo SQFT iProperty Defined for this Part"
'				AddSqft()
'            Err.Clear()
'        End If
'    End Sub
	
		
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
	
	
    Public Sub AddLload()
        Dim ipsetCustom As Inventor.PropertySet
		ipsetCustom = _oDoc.PropertySets.Item("Inventor User Defined Properties")
        Dim iholder As Inventor.Property
		Dim dNum As Double = 0
		iholder = ipsetCustom.Add(dNum, "LLOAD")
		_oDoc.Save 'force update
		_errorState = False 'reset
		'recheck	
        On Error Resume Next
		iLLOAD = ipsetCustom.Item("LLOAD")
        If Err.Number <> 0 Then
            _errorState = True
            _strErrorValue = _strErrorValue & "\\nNo LLOAD iProperty Defined for this Part"
            Err.Clear()
        End If
    End Sub
	
	
'Public Sub AddSqFt()
'        Dim ipsetCustom As Inventor.PropertySet
'		ipsetCustom = _oDoc.PropertySets.Item("Inventor User Defined Properties")
'        Dim iholder As Inventor.Property
'		Dim dNum As Double = 0
'        iholder = ipsetCustom.Add(dNum, "SQFT")
'		_oDoc.Save 'force update
'		_errorState = False 'reset
'		'recheck	
'        On Error Resume Next
'		iSQFT = ipsetCustom.Item("SQFT")
'        If Err.Number <> 0 Then
'            _errorState = True
'            _strErrorValue = _strErrorValue & "\\nNo SQFT iProperty Defined for this Part"
'            Err.Clear()
'        End If
'    End Sub	
	
End Class

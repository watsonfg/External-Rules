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

Public Class RLOAD
    Private _dRLOAD As Double = -999.0
    Private _oDoc As Inventor.PartDocument
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
	Public iRLOAD As Inventor.Property
	
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
				
				Err.Clear
            End If			
			

            iWARNING = iProperties.Item("WARNINGS")
            If Err.Number <> 0 Then
                _errorState = True
                _strErrorValue = _strErrorValue & "\\nNo WARNINGS iProperty Defined for this Part"
				Err.Clear
            End If
			
			
            iRLOAD = iProperties.Item("RLOAD")
            If Err.Number <> 0 Then
                _errorState = True
                _strErrorValue = _strErrorValue & "\\nNo RLOAD iProperty Defined for this Part"
				Err.Clear
            End If	
			

            Dim iHolder As Inventor.Property
            iHolder = iProperties.Item("CorePartMaxCount")

            ' Non standard use of error handling.  Only FACEGROUP parts will contain this iProperty.  If it is a FACEGROUP then mulitple solid bodies are requird to calculate RLOAD
            If Err.Number <> 0 Then
                _bIsFaceGroup = False
				Err.Clear
            Else
                _bIsFaceGroup = True
                _iFaceGroupQty = iHolder.Value
            End If

            Dim invDesignInfo As Inventor.PropertySet
            invDesignInfo = _oDoc.PropertySets.Item("Design Tracking Properties")
            _partName = invDesignInfo.Item("Part Number").Value
        Else
            _errorState = True
            _strErrorValue = "This rule is not being run in an IPT file."
        End If
    End Sub

    ReadOnly Property Value As Double
        Get
            Return _dRLOAD
        End Get
    End Property

    ReadOnly Property CountOfHoles As Integer
        Get
            Return _CountHoles()
        End Get
    End Property
	
	Property IsInErrorState As Boolean
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
	
	ReadOnly Property ErrorValue As String
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
		Dim dTempPerim as Double = 0

        If Not _bIsFaceGroup Then
            On Error Resume Next
            If _oDoc.ComponetDefinition.Features.ExtrudeFeatures(_corePartName).Name = _corePartName Then
                For iLoop = 1 To _oDoc.ComponentDefinition.Features.ExtrudeFeatures(_corePartName).Faces.Count
                    _ShowFacePerimeter(_oDoc.ComponentDefinition.Features.ExtrudeFeatures(_corePartName).Faces(iLoop), dOutWrk, dInWrk)
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
            Else
                _errorState = True
                _strErrorValue = _strErrorValue & "\\nPart is missing a 'CorePart' Extrude Features"
            End If
        Else
            Dim iFGLoop As Integer
            On Error Resume Next
            For iFGLoop = 1 To _iFaceGroupQty
                If _oDoc.ComponetDefinition.Features.ExtrudeFeatures(_corePartName & iFGLoop).IsActive Then
                    For iLoop = 1 To _oDoc.ComponentDefinition.Features.ExtrudeFeatures(_corePartName & iFGLoop).Faces.Count
                        _ShowFacePerimeter(_oDoc.ComponentDefinition.Features.ExtrudeFeatures(_corePartName & iFGLoop).Faces(iLoop), dOutWrk, dInWrk)
                        If dOutWrk >= dOutsideT Then
                            dOutsideT = dOutWrk
                            If (dOutWrk + dInWrk) > dTempPerim Then
                                dTempPerim = dOutWrk + dInWrk
                                _selectedFaceNum = iLoop
                            End If
                        End If
                        If dInWrk >= dInsideT Then
                            dInsideT = dInWrk
                        End If

                        dOutWrk = 0
                        dInWrk = 0
                    Next iLoop
                    dOutside = dOutside + dOutsideT
                    dInside = dInside + dInsideT
                End If
            Next iFGLoop
        End If
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
		
			Select Case iMATTYP.Value
				Case "1PPBK"			'list MATTYP's in descending order for best results!
                    _FeedRate = 400 / 5'Values are Inches per Minute
				Case "1RPPBK"
                    _FeedRate = 400 / 5
				Case "1PPWH"
                    _FeedRate = 400 / 5	
				Case "1LMC"
                    _FeedRate = 400 / 5'waiting for acccurate rate 3-26
				Case "2ACC"
                    _FeedRate = 400 / 5'waiting for acccurate rate 3-26
				Case "2ACS"
                    _FeedRate = 400 / 5'waiting for acccurate rate 3-26
				Case "2ACR"
                    _FeedRate = 400 / 5'waiting for acccurate rate 3-26
				Case "2ACF" 			'PPW-Added 5/6/20
                    _FeedRate = 400 / 5'waiting for acccurate rate 3-26
				Case "2M1"
                    _FeedRate = 400 / 5'waiting for acccurate rate 3-26
                Case "2M2"
                    _FeedRate = 400	/ 5	'Added 8/17/19 CAK EC-1453
				Case "2M1ZE"
                    _FeedRate = 400 / 5'waiting for acccurate rate 2/3/12	
				Case "2MDF"
                    _FeedRate = 400 / 5 'waiting for acccurate rate 3-26
				Case "2TMDF"
                    _FeedRate = 400 / 5	'Added 09.09.2019 DBW
				Case "2TMDF2"
                    _FeedRate = 400 / 5	'Added 8/17/19 CAK EC-1453
				Case "2MP2"
                    _FeedRate = 400 / 5 'waiting for acccurate rate	
                Case "2PHENGS"
                    _FeedRate = 400 / 5
                Case "3MT2"
                    _FeedRate = 400 / 5	'Need Confimation 3/10/2015		
				Case "3PB"
                    _FeedRate = 300 / 5 'Waiting for accurate rate. 2/26/15
				'3TB, 3TB2, 3TN, and 3TNR are no longer options. ECN 15-040.
				'Case "3TB"
                '    _FeedRate = 300 / 5
				'Case "3TB2"
                 '   _FeedRate = 300 / 5 'Need Confirmation 11/19/13
				'Case "3TN"
                '    _FeedRate = 300 / 5 'Need Confirmation 4/14/13		
				'Case "3TNR"
                '    _FeedRate = 300 / 5 'Need Confirmation 4/15/13							
				Case "3ACF"
                    _FeedRate = 400 / 5'waiting for acccurate rate 3-26
				Case "3ACC"	'EC 1889 5/1/20 JTB
                    _FeedRate = 400 / 5'waiting for acccurate rate 3-26
				Case "3PLY"
                    _FeedRate = 300 / 5' NEED CONFIRMATION LMP 10/30/12
				Case "3MP1"
                    _FeedRate = 400	/ 5
				Case "3MP2"
                    _FeedRate = 400	/ 5
				Case "420010"
                    _FeedRate = 400 / 5						
                Case "4M2"
                    _FeedRate = 400	/ 5
				Case "4MP1"
                    _FeedRate = 400	/ 5										'Added by PWC 08.13.2019
				Case "4MP2"
                    _FeedRate = 400	/ 5										'Added by DBW 07.19.2019
				Case "4P2"
                    _FeedRate = 300 / 5
				Case "4PB"
                    _FeedRate = 300 / 5
				Case "4TB"
                    _FeedRate = 300 / 5	
				Case "4PBTB"
                    _FeedRate = 300 / 5	
				Case "4TB2"
                    _FeedRate = 300 / 5	 'Need Confirmation 11/19/13
				Case "4TN"	'New tackboard
                    _FeedRate = 300 / 5	 'Need Confirmation 2/27/15			
				Case "4TNR"	'New tackboard
                    _FeedRate = 300 / 5	 'Need Confirmation 4/15/15		
				Case "4MDF"
                    _FeedRate = 400 / 5	 '	'New MDF Core for Bahn 2.0 screens Added 4/27/17		
                Case "5P1"
                    _FeedRate = 400 / 5					
                Case "5P2"
                    _FeedRate = 400 / 5
                Case "5.5PB"
                    _FeedRate = 300 / 5
				Case "6EB"
                    _FeedRate = 400 / 5	
				Case "6EB1"
                    _FeedRate = 400 / 5			
				Case "6EB2"
                    _FeedRate = 400 / 5
                Case "6M2"
                    _FeedRate = 400 / 5
				Case "6MDF"
                    _FeedRate = 250 / 2 / 5 'Always double pass machinining
				Case "6MP1"
                    _FeedRate = 250 / 2 / 5 'Always double pass machinining
				Case "6MP2"
                    _FeedRate = 250 / 2 / 5 'Always double pass machinining
				Case "6PPWH"
                    _FeedRate = 250 / 2 / 5 'Always double pass machinining	
                Case "6RP"
                    _FeedRate = 400 / 5
                Case "6RP2"
                    _FeedRate = 400 / 5
				Case "8M2"
                    _FeedRate = 400 / 5
				Case "8P1"
                    _FeedRate = 400 / 5
				Case "8P2"
                    _FeedRate = 400 / 5
				Case "8PB"
                    _FeedRate = 400 / 5	
				Case "8ACR"
                    _FeedRate = 250 / 5	
				Case "8DB"
                    _FeedRate = 400 / 5
				Case "8DB2"
                    _FeedRate = 400 / 5	
				Case "8EB"
                    _FeedRate = 400 / 5
				Case "8EB1"
                    _FeedRate = 400 / 5			
				Case "8EB2"
                    _FeedRate = 400 / 5
				Case "9ANY"
                    _FeedRate = 400 / 5
	            Case "9M2"
                    _FeedRate = 400 / 5
                Case "9MP1"
                    _FeedRate = 250 / 2 / 5 'Always double pass machinining
			    Case "9MDF"
                    _FeedRate = 250 / 2 / 5 'Always double pass machinining
				Case "9MDF1"
                    _FeedRate = 250 / 2 / 5 'Always double pass machinining
                Case "9MP2"
                    _FeedRate = 250 / 2 / 5 'Always double pass machinining
                Case "9P1"
                    _FeedRate = 400 / 5
				Case "9P2"
                    _FeedRate = 400 / 5
				Case "9RP"
                    _FeedRate = 400 / 5	
				Case "9RP1"
                    _FeedRate = 400 / 5
				Case "9RP2"
                    _FeedRate = 400 / 5	
				Case "12P1"
                    _FeedRate = 400 / 5	
				Case "12L2"
                    _FeedRate = 400 / 5	
				Case "14P1"
                    _FeedRate = 300 / 5	'waiting for acccurate rate 8/17/2011	
				Case "14P2"
                    _FeedRate = 300 / 5	'waiting for acccurate rate 8/17/2011	
				Case "NAF"
                    _FeedRate = 300 / 5	
				Case "PLEXIC"
                    _FeedRate = 400 / 5'waiting for acccurate rate 3-26	
				Case "S4S"
                    _FeedRate = 300 / 5	
				Case "OSB" 'EC 1744 ADDED FOR OSB FOR SHIPPING BOXES
                    _FeedRate = 400 / 5'waiting for acccurate rate 3-26
                Case Else
                    iWARNING.Value = iWARNING.Value & "\\nNo FeedRate defined for " & iMATTYP.Value & " [I: " & _partName & "]"
					_errorState = True
                    _FeedRate = 1
            End Select
        End If
    End Function

    Private Function _CountHoles() As Integer
        Dim intHoles As Integer = 0

		Dim oHole As Inventor.HoleFeature
        For Each oHole In _oDoc.ComponentDefinition.Features.HoleFeatures
            intHoles = intHoles + oHole.HoleCenterPoints.Count
        Next oHole

        _CountHoles = intHoles
    End Function

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

	Public Sub ShowRLOADFace()
		Call _getPerimeter()
		Dim oFace As Inventor.Face
		oFace = _oDoc.ComponentDefinition.Features.ExtrudeFeatures("CorePart").Faces(_selectedFaceNum)
		' Clear the select set so it doesn't interfere with the highlight.
		_oDoc.SelectSet.Clear
		
		' Set a reference to the UnitsOfMeasure object to use
		' in converting the values obtained to the current
		' document units.  The lengths returned by the API will
		' always be in centimeters.
		Dim oUOM As Inventor.UnitsOfMeasure
		oUOM = _oDoc.UnitsOfMeasure
		
		' Create a string that will contain the loop information.
		Dim strResults As String
		
		' Create the highlight sets.
		Dim oOuterHS As Inventor.HighlightSet
		oOuterHS = _oDoc.CreateHighlightSet
		oOuterHS.Color = _oApp.TransientObjects.CreateColor(255, 0, 0)
		Dim oInnerHS As Inventor.HighlightSet
		oInnerHS = _oDoc.CreateHighlightSet
		oInnerHS.Color = _oApp.TransientObjects.CreateColor(255, 255, 0)
		Dim dMin As Double, dMax As Double
		Dim dLength As Double
		
		' Find the outer loop.
		Dim dOuterLength As Double
		dOuterLength = 0
		Dim oLoop As Inventor.EdgeLoop
		For Each oLoop In oFace.EdgeLoops
			If oLoop.IsOuterEdgeLoop Then
				Dim oEdge As Inventor.Edge
				For Each oEdge In oLoop.Edges
				' Add this edge to the outer highlight set.
					oOuterHS.AddItem(oEdge)
				' Get the length of the current edge.
					Call oEdge.Evaluator.GetParamExtents(dMin, dMax)
					Call oEdge.Evaluator.GetLengthAtParam(dMin, dMax, dLength)
					dOuterLength = dOuterLength + dLength
				Next
				' Add the to the result message string.
				strResults = "Outer Loop Length (red): " & oUOM.GetStringFromValue(dOuterLength, Inventor.UnitsTypeEnum.kDefaultDisplayLengthUnits) & vbCrLf & vbCrLf
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
					' Add this edge to the inner highlight set.
					oInnerHS.AddItem (oEdge)
					' Get the length of the current edge.
					Call oEdge.Evaluator.GetParamExtents(dMin, dMax)
					Call oEdge.Evaluator.GetLengthAtParam(dMin, dMax, dLength)
					dLoopLength = dLoopLength + dLength
				Next
				' Add this loop to the total length.
				dTotalLength = dTotalLength + dLoopLength	
				' Add to the result message string.
				iLoopCount = iLoopCount + 1
				strResults = strResults & "Inner Loop " & iLoopCount & " Length: " & _
						oUOM.GetStringFromValue(dLoopLength, Inventor.UnitsTypeEnum.kDefaultDisplayLengthUnits) & vbCrLf
			End If
		Next	
		' Display the results.
		strResults = strResults & "Total Inner Loop Length (yellow): " & oUOM.GetStringFromValue(dTotalLength, Inventor.UnitsTypeEnum.kDefaultDisplayLengthUnits)
		strResults = strResults & vbCrLf & vbCrLf & "Total Inner and Outer Loop Length: " & oUOM.GetStringFromValue(dTotalLength + dOuterLength, Inventor.UnitsTypeEnum.kDefaultDisplayLengthUnits)
		MsgBox (strResults)
	End Sub
	
	Public Sub FeedRate()
		Dim strResults as String = ""
		
		strResults = "MATTYP:  " & _iProperties.Item("MATTYP").Value & vbCrLf & vbCrLf
		strResults = strResults & "FeedRate:  " & _feedRate()
		
		MsgBox(strResults)
	End Sub
	
    Public Sub Calculate()
		Dim iProperties As Inventor.PropertySet
        iProperties = _oDoc.PropertySets.Item("Inventor User Defined Properties")
        Dim iRLOAD As Inventor.Property
		
		On Error Resume Next
    	iRLOAD = iProperties.Item("RLOAD")
		If Err.Number <> 0 Then
			_errorState = True
            iERROR.Value = iERROR.Value & "\\nNo RLOAD defined [I: " & _partName & "]"
		Else
			_dPerimeter = _getPerimeter() * _constCMtoInch    ' Convert to Inches from cm
			_dRLOAD = _dPerimeter / _FeedRate()
			
			'iRLOAD.Value = _dRLOAD 'Brandt, I commented this original code and added a round below 
			iRLOAD.Value = Round(_dRLOAD,4)
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
	
	
	Public Sub CheckSqft()
        Dim ipsetCustom As Inventor.PropertySet
		ipsetCustom = _oDoc.PropertySets.Item("Inventor User Defined Properties")
        On Error Resume Next
		iSQFT = ipsetCustom.Item("SQFT")
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
				If iEDGSEQ.Value <> "NONE" Then			
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

	Public Sub CheckPremill()
        Dim ipsetCustom As Inventor.PropertySet
		ipsetCustom = _oDoc.PropertySets.Item("Inventor User Defined Properties")
        On Error Resume Next
		iPREMILL = ipsetCustom.Item("PREMILL")
        If Err.Number <> 0 Then
            _errorState = True
            _strErrorValue = _strErrorValue & "\\nNo PREMILL iProperty Defined for this Part"
				AddPremill()
            Err.Clear()
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
		'recheck	
        On Error Resume Next
		iRLOAD = ipsetCustom.Item("RLOAD")
        If Err.Number <> 0 Then
            _errorState = True
            _strErrorValue = _strErrorValue & "\\nNo RLOAD iProperty Defined for this Part"
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
	
	Public Sub AddPremill()
        Dim ipsetCustom As Inventor.PropertySet
		ipsetCustom = _oDoc.PropertySets.Item("Inventor User Defined Properties")
        Dim iholder As Inventor.Property
		Dim strText As String = "External Rload rule run by WoodPart just created this iPoperty, it should get a value from User or WoodAsm rule"
		iholder = ipsetCustom.Add(strText, "PREMILL")
		_oDoc.Save 'force update
		_errorState = False 'reset
		'recheck	
        On Error Resume Next
        iPREMILL = ipsetCustom.Item("PREMILL")
        If Err.Number <> 0 Then
            _errorState = True
            _strErrorValue = _strErrorValue & "\\nNo PREMILL iProperty Defined for this Part"
            Err.Clear()
        End If
    End Sub
	
End Class

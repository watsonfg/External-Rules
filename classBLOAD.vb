' <IsStraightVb>True</IsStraightVb>
Imports System.Diagnostics
'
'  	Initialize by creating new application and calling with ThisApplication
'
'  	Properties
'		Value - (Double) 			Returns current BLOAD value Of Part
'		CountOfHoles - (Integer) 	Returns the Number Of Hole Features in current part
'		IsInErrorState - (Boolean)	Returns True if an Error has been triggered
'		ErrorValue - (String)		Returns error message(s) that generated the error state
'		
'	SubRoutines
'		 (Not Ready)FeedRate()					Displays MessageBox With Feedrate And material type Of current part
'		 Calculate()				Determine the current BLOAD value And populates the BLOAD segment
'		 ShowBLOADFace()			Displays calculation Of BLOAD value And highlight the correct face that Is being used

Public Class BLOAD
    Private _dBLOAD As Double = 42.0
    Private _oDoc As Inventor.PartDocument
    Private _oApp As Inventor.Application
    Private _constCMtoInch As Double = 0.393700787
    Private _errorState As Boolean = False
    Private _strErrorValue As String = ""
    Private _iProperties As Inventor.PropertySet
    Private _partName As String = ""

    Public iWARNING As Inventor.Property
    Public iERROR As Inventor.Property
    Public iBLOAD As Inventor.Property

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


            iBLOAD = iProperties.Item("BLOAD")
            If Err.Number <> 0 Then
                _errorState = True
                _strErrorValue = _strErrorValue & "\\nNo BLOAD iProperty Defined for this Part"
                Err.Clear()
            End If

        Else
            _errorState = True
            _strErrorValue = "This rule is not being run in an IPT file."
        End If
    End Sub

    ReadOnly Property Value() As Double
        Get
            Return _dBLOAD
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

    Public Sub FeedRate()
        Dim strResults As String = ""

        strResults = "MATTYP:  " & _iProperties.Item("MATTYP").Value & vbCrLf & vbCrLf
        ' strResults = strResults & "FeedRate:  " & _feedRate()

        MsgBox(strResults)
    End Sub

    Public Sub Calculate()
        Dim iProperties As Inventor.PropertySet
        iProperties = _oDoc.PropertySets.Item("Inventor User Defined Properties")
        Dim iBLOAD As Inventor.Property

        On Error Resume Next
        iBLOAD = iProperties.Item("BLOAD")
        If Err.Number <> 0 Then
            _errorState = True
            iERROR.Value = iERROR.Value & "\\nNo BLOAD defined [I: " & _partName & "]"
        Else
            iBLOAD.Value = Round(_dBLOAD, 4)
        End If
    End Sub


End Class

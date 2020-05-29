' <IsStraightVb>True</IsStraightVb>
' <FireOthersImmediately>False</FireOthersImmediately>
Imports System.Diagnostics
Imports System.IO.Stream
Imports System.IO
'
'     Initialize by creating new application and calling with ThisApplication
'
'     Properties
'           IsInErrorState - (Boolean)          Returns True if an Error has been triggered
'           ErrorValue - (String)               Returns error message(s) that generated the error state
'           OutPutPath - (String)               Return/set output path for Saved File.
'           SetLogLevel - (Enum LogLevel) 		Return/set log level for object
'			IsCuttubePart - (Boolean)   		Returns True if iProperty IDWEXPORT exists on the part

'
'     SubRoutines
'           CreateDrawing - (void)        		Creates the Drawing for this Part.
'           Log   - (void)                		Write out SysLogD level logging with basic file stream.  Should be updated to LogObject
'           UpdateErrors - (void)         		Update the ERRORS iIsSheetMetalPartProperty
'			


Public Class DrawingExportV2
    Private debugOn As Boolean = False
    Private debugOnOverride As Boolean = False 'Can use 'True' to force debugOn, overrides the enviornment variable on PC running the rule.
    'GH4/7/14 tweaked to only stop suppress feature like EdgeBand* if ipt has PREMILL iProperty with value like 0.059
    Private DoPremill As Boolean = True 'Can use to 'True' to conditionaly stop suppressing feature like EdgeBand* if PREMILL iProperty exists on the ipt with valid value.
    Private InventorVersion As String = "2012" 'Default "2012" or use "2013" to change output file locations.
    Private fileNameExt As String

    Private _oPart As Inventor.PartDocument
    Private _oDoc As Inventor.Document
    Private _oApp As Inventor.Application

    Private _staticDXFpath As String = "\\wfs.local\Watson\Production\LLOAD Files\"
    Private _staticSATpath As String = "\\wfs.local\Watson\Production\RLOAD Files\"
    Private _staticIPTpath As String = "\\wfs.local\Watson\Production\RLOAD Files\IPT_Files\"
	Private _staticJPGpath As String = "JPG\"		' -- BE (4/8/2015)  This is a suffix only  They will be under the RLOAD/LLOAD

    Private _docType As Inventor.DocumentTypeEnum
    Private _errorState As Boolean = False
    Private _strErrorValue As String = ""
    Private _iProperties As Inventor.PropertySet
    Private _lot As String = ""
    Private _outfile As String = ""
    Private _filename As String = ""
    Private _filenameIpt As String = ""
	Private _masterfilepath As String = ""
    Private _masterfilename As String = ""
    Private _outputpath As String = ""
    Private _count As Integer = 0
    Private _edgeFeatureName As String = "EdgeBand*"
    Private _preMilliPropertyName As String = "PREMILL"
    ' Private ex As Exception
    Private ex As String = ""  'BRITTANY
    Private _partName As String = "" 'GH Added 1/23/14


'plan to delete GH 7/9
''''	' ---  These are wrong and need to be removed.
''''    Private _subPathJPEG As String = "" 'GH Added 7/18/14
''''    Private _filenameJPEG As String = "" 'GH Added 7/18/14
''''	'---- End Wrong Stuff
	
    Public iWARNING As Inventor.Property
    Public iERROR As Inventor.Property
    Public iLOT As Inventor.Property
    Public strIDWEXPORT As String = ""
	'GH 10/13/15 tried boolOpenForIDW back False to test disabled background view update setting. IDW was OK but Wood iam didn't got home/fit
	Public boolOpenForIDW As Boolean = False	 '''lmp test 4/7/2016 back to False
	'LMP 4/9/15 forces drawing to open on CADFLOW, (False Runs silent)
	
    'Private Enum LogLevel 'GH commented out per BE to get rule to run
    Enum LogLevel
        Emergency = 70
        Alert = 60
        Critical = 50
        Err = 40
        Warning = 30
        Notice = 20
        Info = 10
        Debug = 0
    End Enum

    Private DefaultLogLevel As LogLevel = LogLevel.Err
    Private CreateIPTOutput As Boolean = False

    Sub New(ByVal ThisDoc As Inventor.Document, ByVal ThisApplication As Inventor.Application)

        Dim ex As Exception
        Dim strEnviro As String = ""
        _docType = ThisDoc.DocumentType
        _errorState = False
        _strErrorValue = ""
        _oApp = ThisApplication
        _oDoc = ThisDoc

        _oApp.SilentOperation = True  'Added because manual attemp to SaveAs was displaying a msg about saving file outside Project location

        ' This code is setting for next DrawingRule update on the Inventor Models.  Would like to pull strEnviro and seperate Live exports from PCM exports
        '         If strEnviro Not Like "*DTA*" Then
        '               _staticDXFpath = "\\wfs.local\watson\Engineering\02_Frontier_Model_Dev\LLOAD Files\" & strEnviro & "\"
        '               consider JPEG export, add if needed
        '               _staticSATpath  = "\\wfs.local\watson\Engineering\02_Frontier_Model_Dev\RLOAD Files\" & strEnviro & "\"
        '         End If
        '

        'GH 1/22/14  temp save untill new below is proven...
        'Below was setting  a Boolean value by 2 other Boolean values with 'And', remember results would be... F = T And F
        '1st (Not sDebug Is Nothing) evaluates to F if there is no ILOGIC_DEBUG... as on newbie modeler PC in Dev...
        'but if its on modeler PC in Dev it goes to T even if value is false = bad!
        '2nd (sDebug = "true") evaluates to T if the ILOGIC_DEBUG is true, as on Server
        'desired results = T for server or modeler PC set to "true" or when debugOnOverride = True
        'Original bad code
        'Dim sDebug As String = Environ("ILOGIC_DEBUG")
        'Me.debugOn = Not sDebug Is Nothing And sDebug = "true"

        'GH New 1/22/14, debugOn is defaulted to False when Dim
        Dim sDebug As String = Environ("ILOGIC_DEBUG")
        If sDebug Is Nothing Then Me.debugOn = True
        If sDebug = "true" Then Me.debugOn = True
        If debugOnOverride = True Then Me.debugOn = True

        If Me.fileNameExt Is Nothing Then
            If Not SharedVariable Is Nothing Then
                If SharedVariable.Exists("A__SERVER_LOG") Then
                    Me.fileNameExt = SharedVariable("A__SERVER_LOG")
                    Log("SharedVariable: Exists", LogLevel.Info)
                End If
            End If
        End If

        If fileNameExt Is Nothing Then
            Me.fileNameExt = "c:\temp\ilogic.log"
            'GH 1/22/14  check if folder exists
            'It did not on several users PC and threw errors and failed because it could not write to log. Also added catch on write tolog
            Dim DevLogFolder As String = "c:\temp" ' test value"c:\tempForceBadPath"
            If Not IO.Directory.Exists(DevLogFolder) Then
                Try 'Create it
                    IO.Directory.CreateDirectory (DevLogFolder)
                    Log("Warning... DevLogFolder did not exist and was created. Path = " & DevLogFolder.ToString(), LogLevel.Info)
                Catch ex
                    UpdateErrors("Catch in path to log file. (See Debug Log).", ex.ToString())
                    Log("DevLogFolder did not exist and could not be created. Attempted path= " & DevLogFolder.ToString(), LogLevel.Err)
                End Try
            End If
        End If


        'GH Commented out this BE code and changed block far below for DefaultLogLevel, to resolve compile errors on 1/20/14
        'SetLogLevel = LogLevel.Err
        DefaultLogLevel = LogLevel.Err

        _iProperties = _oDoc.PropertySets.Item("Inventor User Defined Properties")

        If IsPart() Then
            _oPart = ThisDoc
        End If

        Try
            iERROR = _iProperties.Item("ERRORS")
            Dim CheckERROR As String = CStr(_iProperties.Item("ERRORS").Value)
            'GH Added 1/22/2014 Check for Pre-Existing ERRORS
            If CheckERROR <> "" Then
                Log("***ERROR*** from DrawingExportV2, Exit Sub because ERRORS iProperty had PreExisting value: " & CheckERROR, LogLevel.Err) ' **GH 12/9/13 **
                _errorState = True
            End If
        Catch ex
            Log("***WARNING*** from DrawingExportV2 Catch, created missing ERRORS iProperty for: " & _oDoc.FullFileName, LogLevel.Warning) ' **GH 12/9/13 **
            _iProperties.Add("", "ERRORS")
        End Try

        Try
            iWARNING = _iProperties.Item("WARNINGS")
        Catch ex
            'create missing iProperty and continue
            Log("**WARNING** from DrawingExportV2 Catch, created missing WARNINGS iProperty for: " & _oDoc.FullFileName, LogLevel.Warning) ' **GH 12/9/13 **
            _iProperties.Add("", "WARNINGS")
        End Try

        Try
            _filename = CStr(_iProperties.Item("ERP_LOT").Value)
        Catch ex
            ' Not an error state if it fails.  Optional iProperty
        End Try

        Dim invDesignInfo As Inventor.PropertySet
        invDesignInfo = _oDoc.PropertySets.Item("Design Tracking Properties")
        Try
            _partName = CStr(invDesignInfo.Item("Part Number").Value)
        Catch
            _partName = "NotFound"
            Log("**WARNING** from DrawingExportV2 Catch, could not obtain the _partName, value set to: " & _partName, LogLevel.Warning) ' **GH 12/9/13 **
            ' Not an error state if it fails.  Optional iProperty
        End Try

		'Log ("Within Sub New initialize(), Pre IsSheetMetalPart",LogLevel.Debug)
        If IsSheetMetalPart() Then
            _outputpath = _staticDXFpath
        Else
            _outputpath = _staticSATpath
        End If
		'GH added override for IDWEXPORT or IsCuttubePart, Want Rload even If they are modeled as sheetmetal
		If IsCuttubePart Then
			'Log ("Within read IsCuttubePart(), Test set _outputpath = _staticSATpath for sheetmetal cuttube  ",LogLevel.Debug)
			_outputpath = _staticSATpath
		End If

        Try
            _masterfilepath = IO.Path.GetDirectoryName(_oDoc.FullFileName)
            _masterfilename = IO.Path.GetFileNameWithoutExtension(_oDoc.FullFileName)
        Catch ex
            _errorState = True
            UpdateErrors("Catch in GetFilename. (See Debug Log).", ex.ToString())
            Log("***ERROR*** from DrawingExportV2 Catch In _masterfilepath/name. Cannot Get Path To Filename.", LogLevel.Err) ' **GH 12/9/13 **
            Exit Sub
        End Try

        If _filename = "" Then
			'GH corrected these to remove file extension 5/7/15
			Dim FileNameWithoutExtension As String
			FileNameWithoutExtension = Left(IO.Path.GetFileName(_oDoc.FullFileName), InStrRev(IO.Path.GetFileName(_oDoc.FullFileName),".")-1)
			'Log("*** GH test MyFileNameNoExtension is: " & FileNameWithoutExtension & " , for: " & _partName, LogLevel.Info)
            _filename = "PART_MASTERS\" & FileNameWithoutExtension' was IO.Path.GetFileName(_oDoc.FullFileName)
            _filenameIpt = "IPT_MASTERS\" & FileNameWithoutExtension' was IO.Path.GetFileName(_oDoc.FullFileName)
'plan to delete GH 7/9
''''            _subPathJPEG = "PART_MASTERS\" 'GH Added 7/18/14
''''            _filenameJPEG = FileNameWithoutExtension' was IO.Path.GetFileName(_oDoc.FullFileName) 'GH Added 7/18/14
        Else
            _filenameIpt = _filename
'plan to delete GH 7/9			
''''            _subPathJPEG = "JPEG\" 'GH Added 7/18/14
''''            _filenameJPEG = _filename 'GH Added 7/18/14
        End If
    End Sub

    '['
    'GH Commented out this BE block of code and changed line above for DefaultLogLevel, to resolve compile errors on 1/20/14
    '      Property SetLogLevel() As Enum LogLevel
    '        Get
    '            Return DefaultLogLevel
    '        End Get
    '        Set(ByVal value As Enum LogLevel)
    '            If Not value Then
    '                DefaultLogLevel = value
    '            End If
    '        End Set
    '    End Property
    ']

    Property IsInErrorState() As Boolean
        Get
            Return _errorState
        End Get
        Set(ByVal value As Boolean)
            If Not Value Then
                _errorState = value
                _strErrorValue = ""
            End If
        End Set
    End Property

    ReadOnly Property IsPart() As Boolean
        Get
            Return _oDoc.DocumentType = Inventor.DocumentTypeEnum.kPartDocumentObject
        End Get
    End Property

    ReadOnly Property IsSheetMetalPart() As Boolean
        Get
			Try
				'Log ("Within IsSheetMetalPart(), Within start of Try to get DocumentTypeEnum  ",LogLevel.Debug)
            	Return _oDoc.DocumentType = Inventor.DocumentTypeEnum.kPartDocumentObject AndAlso _
            	_oDoc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
				'Log ("Within IsSheetMetalPart(), End of Try to get IsSheetMetalPart, should return True ",LogLevel.Debug)
			Catch
				' Not an error state if it fails.
				'Log ("Within IsSheetMetalPart(), Within Catch to get IsSheetMetalPart, should return False ",LogLevel.Debug)
			End Try	
        End Get
    End Property
	
	'GH added this sub 7/9/15
	ReadOnly Property IsCuttubePart() As Boolean
		Get
		'Log ("Within IsCuttubePart(), Start Get ",LogLevel.Debug)
			Try
				'Log ("Within IsCuttubePart(), Within start of Try to get IDWEXPORT  ",LogLevel.Debug)
				strIDWEXPORT = CStr(_iProperties.Item("IDWEXPORT").Value)
				IsCuttubePart = True
				'Log ("Within IsCuttubePart(), End of Try to get IDWEXPORT, should return True ",LogLevel.Debug)
			Catch
				' Not an error state if it fails.  Optional iProperty
				IsCuttubePart = False
				'Log ("Within IsCuttubePart(), Within Catch to get IDWEXPORT, should return False ",LogLevel.Debug)
			End Try
			'Log ("Within IsCuttubePart(), After Try to get IDWEXPORT, the IsCuttubePart should return true",LogLevel.Debug)

            Return IsCuttubePart
			'Log ("Within IsCuttubePart(), End Get ",LogLevel.Debug)
        End Get
    End Property

    Property OutputPath() As String
        Get
            Return _outputpath
        End Get
        Set(ByVal value As String)
            If Not Value Then
                _outputpath = value
            End If
        End Set
    End Property

    ReadOnly Property ErrorValue() As String
        Get
            Return _strErrorValue
        End Get
    End Property

    Public Sub CreateDrawing()
		Log ("Enter Public Sub CreateDrawing() ",LogLevel.Debug)
        Dim ex As Exception
		Dim _outputPDF as String
		_outputPDF = _outputpath & "PDF\" & _filename & ".pdf"
        If Not _errorState Then
			
			'Log ("Within Public Sub CreateDrawing, Pre IsCuttubePart",LogLevel.Debug)
			'Find CUTTUBE here, BEFORE sheetmetal or ipt check. GH added 7/9/15
			If IsCuttubePart Then
				'Log ("Within Public Sub CreateDrawing, Within IsCuttubePart",LogLevel.Debug)
				Try
					Call _writeIAMExport()

					Try
						If (CStr(_iProperties.Item("MILL").Value) = "Y") Then
							Call _writeTubeSTP()
						End If
					Catch ex
						' Not an error state if it fails.  Optional iProperty
					End Try
				Catch ex
                    _errorState = True
                    UpdateErrors("Catch in Call _writeIAMExport.", ex.ToString())
                    Log("***ERROR*** in Call _writeIAMExport for " & _oDoc.FullFileName & " Ex= " & ex.ToString(), LogLevel.Err)
                End Try
            ElseIf IsSheetMetalPart Then
				'Log ("Within Public Sub CreateDrawing, Within IsSheetMetalPart",LogLevel.Debug)
                Try
                    Call _writeSheetMetalDXF()
                Catch ex
                    _errorState = True
                    UpdateErrors("Catch in Call _writeSheetMetalDXF.", ex.ToString())
                    Log("***ERROR*** in Call _writeSheetMetalDXF for " & _oDoc.FullFileName & " Ex= " & ex.ToString(), LogLevel.Err)
                End Try

                Try
					oDrawingDocument = _oApp.Documents.Open(_masterfilepath & "\" & _masterfilename & ".idw", boolOpenForIDW) ''' lmp changes to make drawings visable was false 11/24/14
                   	'Dim _outputPDF as String
					'_outputPDF = _outputpath & "PDF\" & _filename & ".pdf"
					If Not System.IO.File.Exists(_outputPDF) Then
                        Try
                            _writeIDW2PDF(oDrawingDocument, _outputPDF)
                        Catch ex
                            _errorState = True
                            UpdateErrors("CreateDrawing() for PDF for: " & _outputPDF & "  ", ex.ToString())
                            Log("***ERROR*** in _writeIDW2PDF Catch on PDF SaveAs: " & _outputPDF & "  Ex=  " & ex.ToString(), LogLevel.Err)
                        End Try
						
                    End If
					Try
                        oDrawingDocument.Close (True)
                    Catch ex
                        Log("***ERROR*** in _writeIAMExport IDW Catch on oDrawingDocument.Close: " & oDrawingDocument.FullFileName, LogLevel.Warning)
                    End Try
					
                Catch ex
                    _errorState = True
                    UpdateErrors("Catch within IsSheetMetalPart trying to open IDW document, Pre Call _writeIDW2PDF." & vbCrLf & _
					"(No IDW found with same name. Could also be a CUTTUBE thats missing IDWEXPORT iProperty.)", ex.ToString())
                    Log("***ERROR*** Catch within IsSheetMetalPart trying to open IDW document, Pre Call _writeIDW2PDF for " & _oDoc.FullFileName & " Ex= " & ex.ToString(), LogLevel.Err)
                End Try
				
				
            ElseIf IsPart Then
				'Log ("Within Public Sub CreateDrawing, Within IsPart",LogLevel.Debug)
                '''GH Testing several ways to force some exceptions
                '               Try
                '                   'Conversion/Comparison
                '                   If "Glen" = True Then' or = 1
                '                       'trying to generate an exception
                '                   End If
                '
                '                   'wrong type value
                '                   '_errorState = "GlenForceError"
                '
                '               Catch ex
                '                   UpdateErrors("Catch in Throw test.", ex.ToString())
                '                   Log("***ERROR*** test Throw... Ex=  "  & ex.ToString(), LogLevel.Err)
                '
                '               End Try
                '                   'Throw New Exception("GhFake New Ex...in IsPart, before Try Call... 3 types of drawing")
                '''end test


                Try
                    'Throw New Exception("GhFake New Ex...in IsPart on Try Call... 3 types of drawing")
					
					'GH 7/9 moved this ABOVE normal wood part exports And added Else so it's ether or not both
					'This allows IDWEXPORT to be used on iam that is not a CUTTUBE, should get same results ether way
                    If strIDWEXPORT = "Y" Then
						Call _writeIAMExport()
					Else
						Call _writeWoodDWG()
                   		Call _writeWoodPDF(_oPart)
                    End If
					
'plan to delete GH 7/9 
'''moved the strIDWEXPORT up ABOVE!!! normal wood part exports And added Else so it's ether or not both					
'''                    Call _writeWoodDWG()
'''                    Call _writeWoodPDF(_oPart)
'''                    If strIDWEXPORT = "Y" Then
'''                        Call _writeIAMExport()
'''                    End If
                Catch ex
                    _errorState = True
                    UpdateErrors("Catch in Call _writeWood(1of3 DWG,PDF,IDW).", ex.ToString())
                    Log("***ERROR*** in Call _writeWood(1of3 DWG,PDF,IDW) for " & _oDoc.FullFileName & "Ex= " & ex.ToString(), LogLevel.Err)
                End Try
            Else
				'Log ("Within Public Sub CreateDrawing, Within 'Else' - assumed iam file",LogLevel.Debug)
				'Assumed iam file here
                Try
					Call _writeIAMExport()
                Catch ex
                    _errorState = True
                    UpdateErrors("Catch in _writeIAMExport.", ex.ToString())
					Log("***ERROR*** in Call _writeIAMExport for " & _oDoc.FullFileName & " Ex= " & ex.ToString(), LogLevel.Err)
                End Try
            End If
        End If

        If Not _strErrorValue = "" Then ' BRITTANY 1/21/14
            Log("***GH test value*** At end of CreateDrawing()...  _strErrorValue:" & _strErrorValue, LogLevel.Info) ' GH 1/21/14
            Log("***GH test value*** At end of CreateDrawing()...  _iProperties.Item('ERRORS').Value:" & CStr(_iProperties.Item("ERRORS").Value), LogLevel.Info) ' GH 1/21/14
        End If
    End Sub



    Private Sub _writeSheetMetalDXF()
        Log(">>>>_writeSheetMetalDXF()", LogLevel.Info)
        Dim ex As Exception
        If Not _errorState Then
            Dim oCompDef As Inventor.SheetMetalComponentDefinition

            Try
                oCompDef = _oDoc.ComponentDefinition
                If oCompDef.HasFlatPattern = False Then
                    Try
                        oCompDef.Unfold()
                    Catch ex
                        Log("***ERROR*** in _writeSheetMetalDXF, Catch on oCompDef.Unfold()" & "  Ex=  " & ex.ToString(), LogLevel.Err)
                    End Try
                Else
                    oCompDef.FlatPattern.Edit()
                End If
            Catch ex
                'GHGH 1/23/14 Note, found frequent catch here, should we Exit sub? Need resolution
                '_errorState = True
                'UpdateErrors("Catch in Get SheetMetalComponentDefinition for: " & _filename, ex.ToString())
                Log("***ERROR*** in _writeSheetMetalDXF Catch on SheetMetalComponentDefinition for " & _oDoc.FullFileName & "  Ex=  " & ex.ToString(), LogLevel.Err)

            End Try
            'Log("***GH test value*** in _writeSheetMetalDXF, _filename: "  & _filename, LogLevel.Err)
            'Log("***GH test value*** in _writeSheetMetalDXF, _oDoc.FullFileName: "  & _oDoc.FullFileName, LogLevel.Err)

            'Build the string that defines the format of the DXF file.
            Dim sOut As String
            'Old original sOut = "FLAT PATTERN DXF?AcadVersion=R12&InvisibleLayers=IV_TANGENT;IV_OUTER_PROFILE;IV_ARC_CENTERS;IV_INTERIOR_PROFILES;IV_BEND;IV_BEND;IV_BEND_DOWN;IV_TOOL_CENTER;IV_TOOL_CENTER_DOWN;IV_TOOL_CENTER;IV_FEATURE_PROFILES;IV_FEATURE_PROFILES;IV_FEATURE_PROFILES_DOWN;IV_ALTREP_FRONT;IV_ALTREP_BACK;IV_UNCONSUMED_SKETCHES;IV_ROLL_TANGENT;IV_ROLL&OuterProfileLayer=0&InteriorProfilesLayer=0"
            'GH commented below line and replaced with new line tested and developed by Fred: To allow Chad to see the center position of punches, I would like to update the drawing rule to not set IV_TOOL_CENTER and IV_TOOL_CENTER_DOWN as invisible layers
            'sOut = "FLAT PATTERN DXF?AcadVersion=R12&InvisibleLayers=IV_TANGENT;IV_OUTER_PROFILE;IV_ARC_CENTERS;IV_INTERIOR_PROFILES;IV_TOOL_CENTER;IV_TOOL_CENTER_DOWN;IV_TOOL_CENTER;IV_FEATURE_PROFILES;IV_FEATURE_PROFILES;IV_FEATURE_PROFILES_DOWN;IV_ALTREP_FRONT;IV_ALTREP_BACK;IV_UNCONSUMED_SKETCHES;IV_ROLL_TANGENT;IV_ROLL&OuterProfileLayer=0&InteriorProfilesLayer=0"
            ' -- * GJD (date) changed line below to show alternate representation of punch tools
            ' -- * GJD (date) sOut = "FLAT PATTERN DXF?AcadVersion=R12&InvisibleLayers=IV_TANGENT;IV_OUTER_PROFILE;IV_ARC_CENTERS;IV_INTERIOR_PROFILES;IV_FEATURE_PROFILES;IV_FEATURE_PROFILES;IV_FEATURE_PROFILES_DOWN;IV_ALTREP_FRONT;IV_ALTREP_BACK;IV_UNCONSUMED_SKETCHES;IV_ROLL_TANGENT;IV_ROLL&OuterProfileLayer=0&InteriorProfilesLayer=0"
            sOut = "FLAT PATTERN DXF?AcadVersion=R12&InvisibleLayers=IV_TANGENT;IV_OUTER_PROFILE;IV_ARC_CENTERS;IV_INTERIOR_PROFILES;IV_FEATURE_PROFILES;IV_FEATURE_PROFILES;IV_FEATURE_PROFILES_DOWN;IV_UNCONSUMED_SKETCHES;IV_ROLL_TANGENT;IV_ROLL&OuterProfileLayer=0&InteriorProfilesLayer=0"
            _outfile = _outputpath & _filename & ".dxf"

            ' Create the DXF file.
            If Not System.IO.File.Exists(_outfile) Then
                Try
                    _oPart.ComponentDefinition.DataIO.WriteDataToFile(sOut, _outfile)
                Catch ex
                    _errorState = True
                    'Check for common errors: Does Path Exist? Is FileName valid
                    If Not System.IO.Directory.Exists(_outputpath) Then 'File.Exists to Directory.Exists 'GH Added 7/18/14
                        UpdateErrors("Catch 'Invalid Path' Error on Export DXF for: " & _outfile & "  ", ex.ToString())
                        Log("***ERROR*** in _writeSheetMetalDXF, Catch on DXF SaveAs, Invalid Path: '" & _outputpath & "'" & "  Ex=  " & ex.ToString(), LogLevel.Err)
                    ElseIf Len(_filename) <= 0 Then
                        UpdateErrors("Catch 'Invalid FileName' Error on Export DXF for: " & _outfile & "  ", ex.ToString())
                        Log("***ERROR*** in _writeSheetMetalDXF, Catch on DXF SaveAs, Invalid : FileName'" & _filename & "'" & "  Ex=  " & ex.ToString(), LogLevel.Err)
                    Else
                        UpdateErrors("Catch in Export DXF for: " & _outfile & "  ", ex.ToString())
                        Log("***ERROR*** in _writeSheetMetalDXF, Catch on DXF SaveAs: " & _outfile & "  Ex=  " & ex.ToString(), LogLevel.Err)
                    End If
                End Try
            End If

            'Create the JPEG file here while view is still Front view of flat pattern. Added by GH on 6/20/14 & 7/18/14
			Try
            	Call _writeSheetMetalJpeg() 'GH Added 7/18/14, GH added Try Catch 7/9/15
			Catch ex
                _errorState = True
                UpdateErrors("Catch in _writeSheetMetalJpeg.", ex.ToString())
				Log("***ERROR*** in Call _writeSheetMetalJpeg for " & _oDoc.FullFileName & " Ex= " & ex.ToString(), LogLevel.Err)
            End Try

				'GH disabled DWFX 9/17/15
'''            ' If there is an idw then Create a DWFX file.
'''            _outputpath = _staticDXFpath 'reset because called _writeSheetMetalJpeg changes the value.  'GH Added 7/18/14
'''            If System.IO.File.Exists(_masterfilepath & "\" & _masterfilename & ".idw") Then
'''                If Not System.IO.File.Exists(_outputpath & _filename & ".dwfx") Then
'''                    Try
'''                        Dim oDrawingDocument As Inventor.Document
'''                        oDrawingDocument = _oApp.Documents.Open(_masterfilepath & "\" & _masterfilename & ".idw", boolOpenForIDW) 'LMP CHANGED 11/24 FROM FALSE NEED TO WORK OUT BETTER SOLUTION
'''                        _outfile = _outputpath & _filename & ".dwfx"
'''                        oDrawingDocument.SaveAs(_outfile, True)
'''                        oDrawingDocument.Close (True)
'''                    Catch ex
'''                        _errorState = True
'''                        UpdateErrors("Catch in Export DWFX for: " & _outfile & "  ", ex.ToString())
'''                        Log("  ***ERROR*** in _writeSheetMetalDXF, Catch on DWFX SaveAs: " & _outfile & "  Ex=  " & ex.ToString(), LogLevel.Err)
'''                    End Try
'''                End If
'''            End If
        End If
        Log("<<<<_writeSheetMetalDXF()", LogLevel.Info)
    End Sub

    Private Sub _writeSheetMetalJpeg() 'GH Added 7/18/14
        Log(">>>>_writeSheetMetalJpeg()", LogLevel.Info)
        Dim ex As Exception

        _outfile = _outputpath & "JPEG\" & _filename & ".jpg" 'JPEG
        'Log(">>>>GH Debug JPEG, _outputpath value is:'" & _outputpath & "'", LogLevel.Debug)
        'Log(">>>>GH Debug JPEG, _outfile value is:'" & _outfile & "'", LogLevel.Debug)

        If Not System.IO.File.Exists(_outfile) Then
            Try
                'Log(">>>>GH Debug JPEG, File did not exist, _outfile value is:'" & _outfile & "'", LogLevel.Debug)
                Dim View As Inventor.View = _oApp.ActiveView
                View.DisplayMode = Inventor.DisplayModeEnum.kWireframeRendering
                'didn't have to change orientation or fit if called just after flat pattern dxf export and before dwfx.
                'results by CADFLow somehow changed to iso and crazy un fit views, yet works fine when run on FATMAN but not in CADFLow, bizzzzare.
                View.Update()
                View.Fit()
                'Log(">>>>Pre Save as in _writeSheetMetalDXF, Catch on JPEG SaveAs, test.", LogLevel.Err)
                Call _oDoc.SaveAs(_outfile, True)
                View.DisplayMode = Inventor.DisplayModeEnum.kShadedWithEdgesRendering 'reset as a convienence when in development doing exports manually
            Catch ex
                _errorState = True
                'Check for common errors: Does Path Exist? Is FileName valid
                If Not System.IO.Directory.Exists(_outputpath & "JPEG\") Then
                    UpdateErrors("Catch 'Directory not found at _outputpath' Error on Export JPEG for: " & _outfile & "  ", ex.ToString())
                    Log("***ERROR*** in _writeSheetMetalDXF, Catch on JPEG SaveAs, Invalid _outputpath Folder: '" & _outputpath & "'" & "  Ex=  " & ex.ToString(), LogLevel.Err)
                ElseIf Len(_filename) <= 0 Then
                    UpdateErrors("Catch '_filename is <=0' Error on Export JPEG for: " & _outfile & "  ", ex.ToString())
                    Log("***ERROR*** in _writeSheetMetalDXF, Catch on JPEG SaveAs, Invalid : FileName'" & _filename & "'" & "  Ex=  " & ex.ToString(), LogLevel.Err)
                Else
                    UpdateErrors("Catch in Export JPEG for: " & _outfile & "  ", ex.ToString())
                    Log("***ERROR*** in _writeSheetMetalDXF, Catch on JPEG SaveAs: " & _outfile & "  Ex=  " & ex.ToString(), LogLevel.Err)
                End If
            End Try
        End If
        Log("<<<<_writeSheetMetalJpeg()", LogLevel.Info)
    End Sub


    Private Sub _writeWoodDWG()
        Log(">>>>_writeWoodDWG()", LogLevel.Info)
        Dim ex As Exception
        Dim _extrudes As New Dictionary(Of String, Boolean)
        Dim _sweeps As New Dictionary(Of String, Boolean)
        Dim strPremillProp As String

        If Not _errorState Then
            Try
                _outfile = _outputpath & _filename & ".dwg"
                'Throw New Exception("GhFake New Ex...in _writeWoodDWG on general Try...")
                If Not File.Exists(_outfile) Then
                    Dim _count As Integer
                    _count = _oPart.ComponentDefinition.Features.ExtrudeFeatures.Count
                    For i As Integer = 1 To _count
                        _extrudes.Add(_oPart.ComponentDefinition.Features.ExtrudeFeatures.Item(i).Name, _oPart.ComponentDefinition.Features.ExtrudeFeatures.Item(i).Suppressed)
                    Next
                    
                    ' Conditionaly suppress edgeband* features, includes quick premill solution.
                    If DoPremill = True Then
                        'if Try finds the PREMILL iProperty then it checks value...
                        'if value is 0.059 or .04 features like edgeband* are Left active so exported DWG Is oversize by EDGTHK
                        'if value is 0.030 (Magna thick tops) or any other value, it suppresses features like edgeband* so parts are cut to finished size.
                        'Log("Test Value to Log, DoPremill In DoPremill _writeWoodDWG() test2: " & _outfile, LogLevel.Err)
                        Try
                            strPremillProp = CStr(_iProperties.Item("PREMILL").Value)
                            'Log("Test Value to Log for strPremillProp is : " & strPremillProp & " In Try _writeWoodDWG() test3: " & _outfile, LogLevel.Err)
                            
                            If strPremillProp = "0.059" Then
                                'leave the feature active
                            ElseIf (strPremillProp = ".04" Or strPremillProp = "0.04" Or strPremillProp = "0.039" ) Then
                                'leave the feature active
                            Else
                                
                                For Each pair As KeyValuePair(Of String, Boolean) In _extrudes
                                    If (pair.Key Like _edgeFeatureName) Then
                                        _oPart.ComponentDefinition.Features.ExtrudeFeatures(pair.Key).Suppressed = True
                                    End If
                                Next
                            End If
                        Catch ex
                            '_errorState = _errorState And False        ' Not an error state if it fails.  Optional iProperty
                                'GH added same suppression here for when PREMILL iProperty is missing
                                'Log("Test Value to Log for Catch of no strPremillProp is : " & strPremillProp & " In Try _writeWoodDWG() test4: " & _outfile, LogLevel.Err)
                                For Each pair As KeyValuePair(Of String, Boolean) In _extrudes
                                    If (pair.Key Like _edgeFeatureName) Then
                                        _oPart.ComponentDefinition.Features.ExtrudeFeatures(pair.Key).Suppressed = True
                                    End If
                                Next
                        End Try
                    Else
                        For Each pair As KeyValuePair(Of String, Boolean) In _extrudes
                            If (pair.Key Like _edgeFeatureName) Then
                                _oPart.ComponentDefinition.Features.ExtrudeFeatures(pair.Key).Suppressed = True
                            End If
                        Next
                    End If

                    ' Supress all sweep features
                    _count = _oPart.ComponentDefinition.Features.SweepFeatures.Count
                    For i As Integer = 1 To _count
                        _sweeps.Add(_oPart.ComponentDefinition.Features.SweepFeatures.Item(i).Name, _oPart.ComponentDefinition.Features.SweepFeatures.Item(i).Suppressed)
                    Next

                    For Each pair As KeyValuePair(Of String, Boolean) In _sweeps
                        _oPart.ComponentDefinition.Features.SweepFeatures(pair.Key).Suppressed = True
                    Next
					
					'GH added 7/15/15 temp fix, really need to evaluate which side of part has most features (like RLOAD rule) and make that side up.
					
					'''ghgh 7/17/15 quick attempt to get dwg to be on home view, failed, later try double open like pdf uses??
'					Log("Temp Test: in _writeWoodDWG() pre Try set view", LogLevel.Info)
'					Try
'						Dim View As Inventor.View = _oApp.ActiveView
'                		View.DisplayMode = Inventor.DisplayModeEnum.kWireframeRendering
'						Log("Temp Test: in _writeWoodDWG() in Try after dim view", LogLevel.Info)
'						View.Update()
'						Log("Temp Test: in _writeWoodDWG() in Try after View.Update", LogLevel.Info)
'						View.GoHome()
'						Log("Temp Test: in _writeWoodDWG() in Try after View.GoHome", LogLevel.Info)
'						View.Fit()
'					Catch ex
'						_errorState = True
'						UpdateErrors("Catch in View Update before Export DWG for: " & _oPart.FullFileName, ex.ToString()) 'want , ex) ---BRITTANY
'						Log("***ERROR*** in View.Update routine, Catch for:" & _oPart.FullFileName & "  Ex=  " & ex.ToString(), LogLevel.Err)
'					End Try
'					Log("Temp Test: in _writeWoodDWG() post Try set view", LogLevel.Info)
					
                    Try
						'Log("Temp Test: in _writeWoodDWG() just before Try SaveAs", LogLevel.Info)
                        _oPart.SaveAs(_outfile, True)
                        'Throw New Exception("GhFake New Ex...in _writeWoodDWG on Try  _oPart.SaveAs...")
                    Catch ex
                        're-check in Catch in case file was made by another session
                        If Not File.Exists(_outfile) Then
                            _errorState = True
                            UpdateErrors("Catch in Export DWG for: " & _outfile & "  ", ex.ToString())
                            Log("***ERROR*** in _writeWoodDWG() DWG Catch on SaveAs: " & _outfile & "  Ex=  " & ex.ToString(), LogLevel.Err)
                        Else
                            Log("File must have been exported by different session, in _writeWoodDWG() DWG Catch on SaveAs: " & _outfile, LogLevel.Info)
                        End If
                    End Try
                End If

                If CreateIPTOutput Then
                    ' Output of .ipt At Current State
                    _outfile = _staticIPTpath & _filenameIpt & ".ipt"
                    If Not File.Exists(_outfile) Then
                        Try
                            _oPart.SaveAs(_outfile, True)
                        Catch ex
                            _errorState = True
                            UpdateErrors("Catch in Export IPT for: " & _outfile & "  ", ex.ToString())
                            Log("***ERROR*** in _writeWoodDWG() IPT Catch on SaveAs: " & _outfile & "  Ex=  " & ex.ToString(), LogLevel.Err)
                        End Try
                    End If
                End If

                ' return all sweeps to previous values
                For Each pair As KeyValuePair(Of String, Boolean) In _sweeps
                    _oPart.ComponentDefinition.Features.SweepFeatures(pair.Key).Suppressed = pair.Value
                Next

                For Each pair As KeyValuePair(Of String, Boolean) In _extrudes
                    _oPart.ComponentDefinition.Features.ExtrudeFeatures(pair.Key).Suppressed = pair.Value
                Next

            Catch ex
                _errorState = True
                UpdateErrors("Catch in _writeWoodDWG . (See Debug Log).", ex.ToString())
                Log("Exception from _writeWoodDWG() General Catch in file: " & _outfile & "  Ex=  " & ex.ToString(), LogLevel.Err)
            End Try
        End If
        Log("<<<<_writeWoodDWG()", LogLevel.Info)
    End Sub



    Private Sub _writeWoodPDF(ByVal oPartDoc As Inventor.PartDocument)
        Log(">>>>_writeWoodPDF()", LogLevel.Info)
        Dim ex As Exception
        If Not _errorState Then
            Dim strExportProp As String
            Dim iProperties As Inventor.PropertySet

            Try
                iProperties = oPartDoc.PropertySets.Item("Inventor User Defined Properties")
                strExportProp = CStr(iProperties.Item("EXPORT").Value)
            Catch ex
                'GH Added 1/23/14 2 checks to see if its a wood part (runs on many ipt for hardware because they have Drawing rule) want to minimize false errors.
                Try
                    strMattypProp = CStr(iProperties.Item("MATTYP").Value)
                    If strMattypProp Like "[569][MP][12]*" Then
                        _errorState = True
                        UpdateErrors("Catch in Get EXPORT iProperty for: " & oPartDoc.FullFileName & "  ", ex.ToString())
                        Log("***ERROR*** Catch in _writeWoodPDF No 'Export' iProperty for:" & oPartDoc.FullFileName & " With MATTYP of: " & strMattypProp, LogLevel.Err)
                        Log("<<<<_writeWoodPDF() -- Hard Exit", LogLevel.Info)
                        Exit Sub
                    Else
                        Try
                            CheckNumericPartNumber = Clng(_partName)
                            'If numeric Skip Warning and let it make PDF
                        Catch ex
                            If Not strMattypProp = "" Then
                                Log("***WARNING*** Catch in _writeWoodPDF No 'Export' iProperty for:" & oPartDoc.FullFileName & " With MATTYP of: " & strMattypProp & " , and with Non-Numeric Part Number of: " & _partName, LogLevel.Warning)
                            Else
                                Log("***WARNING*** Catch in _writeWoodPDF No 'Export' iProperty for:" & oPartDoc.FullFileName & " With no MATTYP, and with Non-Numeric Part Number of: " & _partName, LogLevel.Warning)
                            End If
                        End Try
                    End If
                Catch ex
                    Try
                        CheckNumericPartNumber = Clng(_partName)
                        'If numeric Skip Warning and let it make PDF
                    Catch ex
                        If strMattypProp = "" Then
                            Log("***WARNING*** Catch in _writeWoodPDF No 'Export' iProperty for:" & oPartDoc.FullFileName & " With no MATTYP, and with Non-Numeric Part Number of: " & _partName, LogLevel.Warning)
                        Else
                            Log("***WARNING*** Catch in _writeWoodPDF No 'Export' iProperty for:" & oPartDoc.FullFileName & " With MATTYP of: " & strMattypProp & " , and with Non-Numeric Part Number of: " & _partName, LogLevel.Warning)
                        End If
                    End Try
                End Try
            End Try
			
            _outfile = _outputpath & "PDF\" & _filename & ".pdf" 'was _outfile = _outputpath & _filename & ".pdf" updated by GH 5/7/15
			'Log("*** GH test _outfile is: " & _outfile & " , for: " & _partName, LogLevel.Info)
            If Not System.IO.File.Exists(_outfile) Then
                For Each sketchX As Inventor.PlanarSketch In oPartDoc.ComponentDefinition.Sketches
                    If sketchX.Name = strExportProp Then
                        sketchX.Visible = True
                        sketchX.DimensionsVisible = True
                    End If
                Next

                '''re-open part to force it to ActiveView, hurray!! it works!!
                Dim j2 As Inventor.PartDocument
                j2 = _oApp.Documents.Open(oPartDoc.FullFileName, boolOpenForIDW) '''lmp was true made it boolOpenForIDW
                Dim View As Inventor.View = _oApp.ActiveView
				'GH 10/13/15 test next line
				'Inventor.ActiveView.Fit

                '''leave this for a while just to be sure...
                '                       '''Dim a default then select which view type gives most consistent results...
                '                       ''Probably a better way but this works!! 3 or 4 times for iam then failed, worked OK several times after restart of App
                '                       ''found it exports on non existing view(corrupt) when the active view is not present
                '                       Dim View As Inventor.View = _oApp.Views(1)'This was original code, (1) was to specify first tab **GH 12/9/13 **
                '                       Log("   GH test value for:" & oPartDoc.FullFileName & "  ViewType is: " & ViewType) ' **GH 12/9/13 **
                '
                '                       If ViewType = "ipt" Then
                '                             View = _oApp.ActiveView' **GH 12/9/13 **
                '                       ElseIf ViewType = "iam" Then
                '                       'try re-open doc here to force as active??
                '                             'use default _oApp.Views(1) set above but log an error
                '                             'View = _oApp.Views(1)'This was original code, (1) was to specify first tab **GH 12/9/13 **
                '                       Else
                '                             'use default _oApp.Views(1) set above but log an error
                '                             Log("***ERROR*** Unknown in ViewType for:" & oPartDoc.FullFileName) ' **GH 12/9/13 **
                '                       End If

                Try
                    View.DisplayMode = Inventor.DisplayModeEnum.kWireframeRendering
                    View.Update()
                    View.GoHome()
                    View.Fit()
                Catch ex
                    _errorState = True
                    UpdateErrors("Catch in View Update before Export PDF for: " & oPartDoc.FullFileName, ex.ToString()) 'want , ex) ---BRITTANY
                    Log("***ERROR*** in View.Update routine, Catch for:" & oPartDoc.FullFileName & "  Ex=  " & ex.ToString(), LogLevel.Err)
                End Try

                Try
                    oPartDoc.SaveAs(_outfile, True)
                Catch ex
                    're-check in Catch in case file was made by another session
                    If Not File.Exists(_outfile) Then
                        _errorState = True
                        UpdateErrors("Catch in Export PDF for: " & oPartDoc.FullFileName & "  ", ex.ToString())
                        Log("***ERROR*** in _writeWoodPDF Catch on PDF SaveAs: " & _outfile & "  Ex=  " & ex.ToString(), LogLevel.Err)
                    Else
                        Log("File must have been exported by different session, in _writeWoodPDF() PDF Catch on SaveAs: " & _outfile & "  Ex=  " & ex.ToString(), LogLevel.Err)
                    End If
                End Try

                'View.DisplayMode = Inventor.DisplayModeEnum.kShadedRendering'GHchanged 1/28/15
                View.DisplayMode = Inventor.DisplayModeEnum.kShadedWithEdgesRendering
                View.Update()

                ' Turn off Sketch
                For Each sketchX As Inventor.PlanarSketch In oPartDoc.ComponentDefinition.Sketches
                    If sketchX.Name = strExportProp Then
                        sketchX.Visible = False
                    End If
                Next
            End If
        End If
        Log("<<<<_writeWoodPDF()", LogLevel.Info)
    End Sub

	Private Sub _writeIDW2PDF(odrawing as Inventor.Document, strPath As String)
	' Get the PDF translator Add-In.
        Dim oPDFTrans As Inventor.TranslatorAddIn
        oPDFTrans = _oApp.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")
        If oPDFTrans Is Nothing Then
            Log( "Could not access PDF translator.", LogLevel.Err)
            Exit Sub
        End If
        
        ' Create some objects that are used to pass information to
        'the translator Add-In.

		Dim oContext As Inventor.TranslationContext
		oContext = _oApp.TransientObjects.CreateTranslationContext
		Dim oOptions As Inventor.NameValueMap
		oOptions = _oApp.TransientObjects.CreateNameValueMap
		If oPDFTrans.HasSaveCopyAsOptions(odrawing, _oContext, oOptions) Then
			'GH added the following 3 options 8/11/15 set both check boxes to True (see them on Options tab in SaveAs PDF) and increased DPI from 400 to 600
			Log ("Within _writeIDW2PDF, Pre set Export Options.",LogLevel.Debug)
			oOptions.Value("All_Color_AS_Black") = 1
			oOptions.Value("Remove_Line_Weights") = 1
			oOptions.Value("Vector_Resolution") = 600
			' Set to print all sheets.  This can also have the value
			' kPrintCurrentSheet or kPrintSheetRange. If kPrintSheetRange
			' is used then you must also use the CustomBeginSheet and
			' Custom_End_Sheet to define the sheet range.
			oOptions.Value("Sheet_Range") = Inventor.PrintRangeEnum.kPrintAllSheets
		
			' Other possible options...
			'oOptions.Value("Custom_Begin_Sheet") = 1
			'oOptions.Value("Custom_End_Sheet") = 5
			'oOptions.Value("All_Color_AS_Black") = True
			'oOptions.Value("Remove_Line_Weights") = True
			'oOptions.Value("Vector_Resolution") = 200
		
			' Define various settings and input to provide the translator.
			oContext.Type = Inventor.IOMechanismEnum.kFileBrowseIOMechanism
			Dim oData As Inventor.DataMedium
			oData = _oApp.TransientObjects.CreateDataMedium()
			oData.FileName = strPath
		
			' Call the translator.
			Call oPDFTrans.SaveCopyAs(odrawing, oContext, oOptions, oData)
		End If
		            
        oPDFTrans = Nothing
        
	End Sub

    Private Sub _writeIAMExport()
        Log(">>>>_writeIAMExport()", LogLevel.Info)
        Dim ex As Exception
        If Not _errorState Then 'GH Added 1/23/14
            If System.IO.File.Exists(_masterfilepath & "\" & _masterfilename & ".idw") Then
'''			'gh9/17BelowLogOn
'''			Log ("Within _writeIAMExport(), After checking, found existing IDW named: " & _masterfilepath & "\" & _masterfilename & ".idw" ,LogLevel.Debug)
'''				'gh9/17Below should be removed or tweaked so DWFX are not required in any logic
'''                If Not System.IO.File.Exists(_masterfilepath & "\" & _masterfilename & ".dwfx") Then
                    Dim oDrawingDocument As Inventor.Document
'''					'gh9/17BelowLogOn
'''					Log ("Within _writeIAMExport, (No DWFX exists) before setting the oDrawingDocument",LogLevel.Debug)
                    oDrawingDocument = _oApp.Documents.Open(_masterfilepath & "\" & _masterfilename & ".idw", boolOpenForIDW) ''' lmp changes to make drawings visable was false 11/24/14
					
		' Added for idw -> PDF
					Dim _outputPDF as String
					_outputPDF = _outputpath & "PDF\" & _filename & ".pdf"
					Log ("Within _writeIAMExport, The _outputPDF is:" & _outputPDF,LogLevel.Debug)
					If Not System.IO.File.Exists(_outputPDF) Then
                        Try
							Log ("Within _writeIAMExport, In Try _writeIDW2PDF, The _outputPDF is:" & _outputPDF,LogLevel.Debug)
                            _writeIDW2PDF(oDrawingDocument, _outputPDF)
                        Catch ex
                            _errorState = True
                            UpdateErrors("Catch in Export PDF for: " & _outputPDF & "  ", ex.ToString())
                            Log("***ERROR*** in _writeIAMExport Catch on _writeIDW2PDF SaveAs: " & _outputPDF & "  Ex=  " & ex.ToString(), LogLevel.Err)
                        End Try
                    End If

					Try
                        oDrawingDocument.Close (True)
                    Catch ex
                        Log("***ERROR*** in _writeIAMExport IDW Catch on oDrawingDocument.Close: " & oDrawingDocument.FullFileName, LogLevel.Warning)
                    End Try
'''				'gh9/17BelowAddedElse
'''				Else
'''					'gh9/17BelowAddedLogOn
'''					Log ("Within _writeIAMExport(), In Else, must have found existing DWFX named: " & _masterfilepath & "\" & _masterfilename & ".dwfx" ,LogLevel.Debug)
'''             End If
            Else
				'No IDW file found...

            	'Log ("Within _writeIAMExport(), After checking, didn't find an IDW named: " & _masterfilepath & "\" & _masterfilename & ".idw" ,LogLevel.Debug)
				' Declare a variable to handle a reference to a document.
                Dim iProperties As Inventor.PropertySet
                Dim i As Integer

                Dim oRefDoc As Inventor.Document
                Dim boolhit As Boolean = False
                Dim DoWoodPDF As Boolean = True
                Dim j As Inventor.PartDocument = Nothing
				
				'If CUTTUBE create simple PDF of part
				If IsCuttubePart Then
					DoWoodPDF = False
					'Log ("Within _writeIAMExport, No IDW found for IsCuttubePart = True",LogLevel.Debug)
					Try
						'Log ("Within _writeIAMExport, No IDW found for IsCuttubePart, Pre Call _writeWoodPDF",LogLevel.Debug)
						'below creates the simple view of part without IDW but was in wrong location, was in LLOAD, expecting RLOAD
						'could use this to change path???
'						Log ("Within _writeIAMExport, Temp test1of3 BEFORE update... Vlaue for _outputpath: = " & _outputpath & "",LogLevel.Debug)
'						_outputpath = _staticDXFpath
'						Log ("Within _writeIAMExport, Temp test2of3 AFTER update... Vlaue for _outputpath: = " & _outputpath & "",LogLevel.Debug)
						'got this far but WoodPDF fails, looking for special Export iProperty to define which sketch to make visible.
'						Call _writeWoodPDF(_oPart)
'						Log ("Within _writeIAMExport, Temp test3of3 Post Call _writeWoodPDF... Vlaue for _outputpath: = " & _outputpath & "",LogLevel.Debug)
					Catch ex
						_errorState = True
						UpdateErrors("Catch in Call _writeIAMExport for No IDW found for IsCuttubePart.", ex.ToString())
						Log("***ERROR*** in Call _writeIAMExport for No IDW found for IsCuttubePart. " & " Ex= " & ex.ToString(), LogLevel.Err)
					End Try
				End If
				
				'normal wood part export for iam
				'Log ("Within _writeIAMExport, No IDW found, Between IsCuttubePart and DoWoodPDF, Vlaue for DoWoodPDF: = " & DoWoodPDF & "",LogLevel.Debug)
				'Log ("Within _writeIAMExport, No IDW found, Between IsCuttubePart and DoWoodPDF, Vlaue for IsCuttubePart: = " & IsCuttubePart & "",LogLevel.Debug)
				If DoWoodPDF = True Then
					' Look at all the open documents in iam looking for ipt with RLOAD iProperty
					For Each oRefDoc In _oDoc.ReferencedDocuments
						If oRefDoc.DocumentType = Inventor.DocumentTypeEnum.kPartDocumentObject Then
							iProperties = oRefDoc.PropertySets.Item("Inventor User Defined Properties")
							i = 1
							Do While i <= iProperties.Count
								If iProperties.Item(i).Name = "RLOAD" Then
									'Log("***GH test value*** in loop of ref ipt with RLOAD...  oRefDoc.FullFileName:" & oRefDoc.FullFileName & " ", LogLevel.Debug)
									If Not boolhit Then
										boolhit = True
										j = _oApp.Documents.Open(oRefDoc.FullFileName, boolOpenForIDW)''' LMP was true now boolOpenForIDW
										j.Close (False) 'close it here then re-open after call, provides more reliable active view
										Call _writeWoodPDF(j)
										Exit Do
									Else
										'must be a second ipt with RLOAD found, Log and Exit For
										Log("***ERROR*** 1 of 2 in _writeIAMExport Catch on multiple ipt with Rload found in iam: " & _oDoc.FullFileName & " ", LogLevel.Err)
										Log("***ERROR*** 2 of 3 1st Rload ipt is: " & j.FullFileName, LogLevel.Err)
										Log("***ERROR*** 3 of 3 2nd Rload ipt is: " & oRefDoc.FullFileName, LogLevel.Err)
										Log("***ERROR*** 4 of 4 Must not be simple wood asm, probably need an idw for asm: ", LogLevel.Err)
										Exit For
									End If
								End If
								i = i + 1
							Loop
						End If
					Next
				End If
				
            End If
        End If
        Log("<<<<_writeIAMExport()", LogLevel.Info)
    End Sub

    Private Sub _writeTubeSTP() 'GH Added 7/18/14
        Log(">>>>_writeTubeSTP()", LogLevel.Info)
        Dim ex As Exception

        _outfile = _outputpath & "STP\" & _filename & ".stp" 'STEP file

        If Not System.IO.File.Exists(_outfile) Then
            Try
                Dim View As Inventor.View = _oApp.ActiveView
                'View.DisplayMode = Inventor.DisplayModeEnum.kWireframeRendering
                'didn't have to change orientation or fit if called just after flat pattern dxf export and before dwfx.
                'results by CADFLow somehow changed to iso and crazy un fit views, yet works fine when run on FATMAN but not in CADFLow, bizzzzare.
                'View.Update()
                'View.Fit()
                'Log(">>>>Pre Save as in _writeSheetMetalDXF, Catch on JPEG SaveAs, test.", LogLevel.Err)
                Call _oDoc.SaveAs(_outfile, True)
                'View.DisplayMode = Inventor.DisplayModeEnum.kShadedWithEdgesRendering 'reset as a convienence when in development doing exports manually
            Catch ex
                _errorState = True
                'Check for common errors: Does Path Exist? Is FileName valid
                If Not System.IO.Directory.Exists(_outputpath & "STP\") Then
                    UpdateErrors("Catch 'Directory not found at _outputpath' Error on Export STP for: " & _outfile & "  ", ex.ToString())
                    Log("***ERROR*** in _writeTubeSTP, Catch on STP SaveAs, Invalid _outputpath Folder: '" & _outputpath & "'" & "  Ex=  " & ex.ToString(), LogLevel.Err)
                ElseIf Len(_filename) <= 0 Then
                    UpdateErrors("Catch '_filename is <=0' Error on Export STP for: " & _outfile & "  ", ex.ToString())
                    Log("***ERROR*** in _writeTubeSTP, Catch on STP SaveAs, Invalid : FileName'" & _filename & "'" & "  Ex=  " & ex.ToString(), LogLevel.Err)
                Else
                    UpdateErrors("Catch in Export STP for: " & _outfile & "  ", ex.ToString())
                    Log("***ERROR*** in _writeTubeSTP, Catch on STP SaveAs: " & _outfile & "  Ex=  " & ex.ToString(), LogLevel.Err)
                End If
            End Try
        End If
        Log("<<<<_writeTubeSTP()", LogLevel.Info)
    End Sub

    Sub Log(ByVal msg As String, ByVal Level As Integer)
        Dim ex As Exception
        Level = 40
        If Me.debugOn Then
            Level = DefaultLogLevel
            'Log("***Test*** Level:" & Level, LogLevel.Err) ' **GH 12/9/13 **
        End If

        If DefaultLogLevel = LogLevel.Debug Then
            Try
                Dim fs As New System.IO.FileStream(fileNameExt & ".debug", System.IO.FileMode.Append, System.IO.FileAccess.Write)
                Dim theLog As New System.IO.StreamWriter(fs)
                theLog.WriteLine (DateAndTime.Now().ToString() & "   " & msg)
                theLog.Close()
            Catch ex
                'May change to Warning later, for now want to know when there are issues writing to log GH 1/22/14
                UpdateErrors("Catch in Try Log.debug. Could not write to fileNameExt = " & fileNameExt & "  ", ex.ToString())
            End Try

        ElseIf Level >= DefaultLogLevel Then
            Try
                Dim fs As New System.IO.FileStream(fileNameExt, System.IO.FileMode.Append, System.IO.FileAccess.Write)
                Dim theLog As New System.IO.StreamWriter(fs)
                theLog.WriteLine (DateAndTime.Now().ToString() & "   " & msg)
                'theLog.WriteLine(DateAndTime.Now().ToString() & " **Test in Sub Log, Level: " & Level)
                theLog.Close()
            Catch ex
                'May change to Warning later, for now want to know when there are issues writing to log GH 1/22/14
                UpdateErrors("Catch in Try Log. Could not write to fileNameExt = " & fileNameExt & "  ", ex.ToString())
            End Try
        End If
    End Sub

    'BRITTANY - changed MyEx to string, send ex.ToString() when calling UpdateErrors
    Sub UpdateErrors(ByVal msg As String, ByVal MyEx As String)
        Dim ErrorMsg As String = ""
        If msg <> "" Then
            'Build full ErrorMsg string, with or without Exception, must be <=503 to populate iProperty
            If Not MyEx <> "" Then
                ErrorMsg = Left(Trim(msg) & "  Ex=" & Trim(MyEx), 503)
            Else
                ErrorMsg = Left(Trim(msg), 503)
            End If

            If CStr(_iProperties.Item("ERRORS").Value) <> "" Then
                'Check if current value is same as what would be added
                If CStr(_iProperties.Item("ERRORS").Value) <> ErrorMsg Then
                    'Combine existing value and new msg, want as much err info returned as possible right now
                    _iProperties.Item("ERRORS").Value = Left(CStr(_iProperties.Item("ERRORS").Value) & "~~" & ErrorMsg, 503)
                End If
            Else
                _iProperties.Item("ERRORS").Value = ErrorMsg
            End If
        End If
    End Sub

    '************************************************************************************************************************************
    '*      1/16/2014   (BE) Clean up and revision of Class.
    '*                              % Commited all tested changes added to date.
    '*                              % Improved Log fuction to act as a simple SyslogD logger so mulitiple levels of logging could be turned on or
    '*                                turned off through functions and controls.
    '*                              % More Robust error checking.  Class is not assuming that any error checking is being done by the Calling
    '*                                entity
    '*                              % Added control for PREMILL
    '************************************************************************************************************************************
    '*      1/22/2014   (GH&BRITTANY) Tested and debugged the 1/16/2014 Improved Log fuctions
    '*                              % Added T/F Toggle for PREMILL
    '*                              % Added Inventor 2013 version Toggle so 1 universal rule can easily be run everywhere
    '*                              % Tested and improved caught errors to both Log and ERRORS iProperty
    '*                              % Added check for Path to development log
    '*                              % Added Try/Catch to Write log, returns ERRORS
    '************************************************************************************************************************************
    '*      4/7/2014   (GH) Tweaked the code for PREMILL, nearly reveresed original logic
    '*                              % Tested only rule change
    '*                              % Suspect Log Level routine logic is backwards, tried Info & Debug, got nothing, get Info only when set to Err.
    '************************************************************************************************************************************
    '*      7/??/2014   (LMP) Tweaked the code for PREMILL, added .04 / 0.04
    '*      7/??/2014   (GJD) Tweaked the sOUTcode for something with punch
    '************************************************************************************************************************************
    '*      7/18/2014   (GH) Added code for JPEG export of all sheetmetal parts
    '*                              % Had trouble with view as iso, fixed by running right after dxf and before dwfx
    '*                              % dwfx failed, more testing found path was tweaked and had to be reset
    '************************************************************************************************************************************
    '*  7/9 to 7/17   2015   (GH) Added code for corrected CUTTUBE export based on IDWEXPORT to find IsCuttubePart
    '*                              % Removed DWFX
    '*                              % Cleaned up old comments, removed unused code including all: _subPathJPEG and _filenameJPEG
	'*                             % Tried change DWG to View.Home but it made no difference, commented it out
    '************************************************************************************************************************************
	 '*  9/17   2015   (GH) Commented out code for DWFX export 
    '*                              % Added some temp log debugs to track down why DWFX are still occuring.
    '*                              % 
    '************************************************************************************************************************************
	
	
	

End Class







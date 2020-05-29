' <IsStraightVb>True</IsStraightVb>
Imports System.Diagnostics
Imports System.IO.Stream
Imports System.IO
'
'  	Initialize by creating new application and calling with ThisApplication
'
'  	Properties
'		IsInErrorState - (Boolean)	Returns True if an Error has been triggered
'		ErrorValue - (String)		Returns error message(s) that generated the error state
'		OutPutPath - (String)		Return/set output path for Saved File.
'		
'	SubRoutines
'       CreateDrawing - (void)      Creates the Drawing for this Part.
'	

Public Class WatsonUtils
    Private _oPart As Inventor.PartDocument
    Private _oDoc As Inventor.Document
    Private _oApp As Inventor.Application

    Sub New(ByVal ThisDoc As Inventor.Document, ByVal ThisApplication As Inventor.Application)

        Dim ex As Exception
        Dim strEnviro As String = ""
        _docType = ThisDoc.DocumentType
        _errorState = False
        _strErrorValue = ""
        _oApp = ThisApplication
        _oDoc = ThisDoc

        If IsPart() Then
            _oPart = ThisDoc
        End If

    End Sub

    ReadOnly Property IsPart() As Boolean
        Get
            Return _oDoc.DocumentType = Inventor.DocumentTypeEnum.kPartDocumentObject
        End Get
    End Property

    ReadOnly Property IsSheetMetalPart() As Boolean
        Get
            Return _oDoc.DocumentType = Inventor.DocumentTypeEnum.kPartDocumentObject AndAlso _
            _oDoc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
        End Get
    End Property

    Public Sub CleaniLogic()

        ' Declare a variable to handle a reference to a document.
        Dim iProperties As Inventor.PropertySet
        Dim i As Integer

        Dim oRefDoc As Inventor.Document
        ' Look at all the open documents

        For Each oRefDoc In _oDoc.ReferencedDocuments
            If oRefDoc.DocumentType = Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
                If oRefDoc.ComponentDefinition.RepresentationsManager.ActiveLevelOfDetailRepresentation.Name <> "iLogic" Then
                    Dim oLOD As Inventor.AssemblyDocument.ComponentDefinition.RepresentationsManager.LevelOfDetailRepresentations
                    Dim oAsmCompDef AsInventor.AssemblyDocument.ComponentDefinition
                    oAsmCompDef = oRefDoc.ComponentDefinition
                    oLOD = oAsmCompDef.RepresentationsManager.LevelOfDetailRepresentations.Item("iLogic")
                    oLOD.Activate(True)
                End If
            End If
        Next


    End Sub



End Class

'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Text
Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Arguments

  ''' <summary>
  ''' Base Class
  ''' </summary>
  ''' <remarks></remarks>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class PlugInExecuteReturnArgs

#Region "Class Variables"

    Private mblnReturnCode As PlugInExecuteReturnCode
    Private mobjException As Exception
    Private mobjDocument As Document
    Private mobjRecord As Object
    Private mstrDocumentProcessedPath As String = String.Empty
    Private mstrFolderToFileIn As String = String.Empty

#End Region

#Region "Public Enumerations"

    Public Enum PlugInExecuteReturnCode
      Failure = 0     'PlugIn had a failure, tell the caller to fail the document
      Success = 1     'PlugIn Succeeded 
      DontProcess = 2 'End the processing of the document
      SkipEventExecution = 3 'Skip executing the specific event
    End Enum

#End Region

#Region "Constructors"

    Public Sub New()
      Try

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpReturnCode As PlugInExecuteReturnCode)
      Try
        ReturnCode = lpReturnCode
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpReturnCode As PlugInExecuteReturnCode,
                   ByVal lpDocument As Document)
      Try
        ReturnCode = lpReturnCode
        Document = lpDocument
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpReturnCode As Boolean,
                   ByVal lpDocumentProcessedPath As String,
                   ByVal lpDocument As Document)
      Me.New(lpReturnCode, lpDocumentProcessedPath, lpDocument, Nothing, Nothing)
    End Sub

    Public Sub New(ByVal lpReturnCode As Boolean,
                   ByVal lpDocumentProcessedPath As String,
                   ByVal lpDocument As Document,
                   ByVal lpRecord As Object,
                   ByVal lpException As Exception)
      Try
        ReturnCode = lpReturnCode
        DocumentProcessedPath = lpDocumentProcessedPath
        Document = lpDocument
        Record = lpRecord
        Exception = lpException
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' The original path of the file to be processed
    ''' This is the same path passed into PlugInExecuteDocProcessingEventArgs
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DocumentProcessedPath() As String
      Get
        Return mstrDocumentProcessedPath
      End Get
      Set(ByVal value As String)
        mstrDocumentProcessedPath = value
      End Set
    End Property

    ''' <summary>
    ''' The resulting document after DocumentToProcessPath has been processed
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Document() As Document
      Get
        Return mobjDocument
      End Get
      Set(ByVal value As Document)
        mobjDocument = value
      End Set
    End Property

    Public Property ReturnCode() As PlugInExecuteReturnCode
      Get
        Return mblnReturnCode
      End Get
      Set(ByVal value As PlugInExecuteReturnCode)
        mblnReturnCode = value
      End Set
    End Property

    Public Property Exception() As Exception
      Get
        Return mobjException
      End Get
      Set(ByVal value As Exception)
        mobjException = value
      End Set
    End Property

    Public Property Record() As Object
      Get
        Return mobjRecord
      End Get
      Set(ByVal value As Object)
        mobjRecord = value
      End Set
    End Property

    Public Property FolderToFileIn() As String
      Get
        Return mstrFolderToFileIn
      End Get
      Set(ByVal value As String)
        mstrFolderToFileIn = value
      End Set
    End Property

#End Region

#Region "Private Methods"

    Protected Friend Overridable Function DebuggerIdentifier() As String
      Try
        Dim lobjReturnBuilder As New StringBuilder

        ' Append the return code
        lobjReturnBuilder.AppendFormat("{0}; ", ReturnCode.ToString)

        ' Append the document processed path
        If DocumentProcessedPath.Length > 0 Then
          lobjReturnBuilder.AppendFormat("DocPath={0}; ", DocumentProcessedPath)
        End If

        ' Append the document
        If Document IsNot Nothing Then
          lobjReturnBuilder.AppendFormat("{0}; ", Document.DebuggerIdentifier)
        End If

        If Exception IsNot Nothing Then
          lobjReturnBuilder.AppendFormat("{0}:{1}; ", Exception.GetType.Name, Exception.Message)
        End If

        Return lobjReturnBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace

'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports System.Text
Imports Documents.Core
Imports Documents.Migrations
Imports Documents.Utilities

#End Region

Namespace Arguments

  ''' <summary>
  ''' Contains all the parameters necessary for the ImportDocument method.
  ''' </summary>
  ''' <remarks>Used as a sole parameter for the ImportDocument method.</remarks>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class ImportDocumentArgs
    Inherits BackgroundWorkerEventArgs

#Region "Class Variables"

    Private mobjDocument As Document
    Private menuFilingMode As FilingMode
    Private mobjPathFactory As PathFactory
    Private mblnImportAsPackage As Boolean = True
    Private mblnSetAnnotations As Boolean = True
    Private mblnSetPermissions As Boolean = True
    Private mstrFolderToFileIn As String = String.Empty
    Private menuVersionType As VersionTypeEnum = VersionTypeEnum.Unspecified

#End Region

#Region "Public Properties"

    Public Property Document() As Document
      Get
        Return mobjDocument
      End Get
      Set(ByVal value As Document)
        mobjDocument = value
      End Set
    End Property

    Public ReadOnly Property FilingMode() As FilingMode
      Get
        Return menuFilingMode
      End Get
    End Property

    Public Property ImportAsPackage() As Boolean
      Get
        Return mblnImportAsPackage
      End Get
      Set(ByVal value As Boolean)
        mblnImportAsPackage = value
      End Set
    End Property

    Public ReadOnly Property PathFactory() As PathFactory
      Get
        Return mobjPathFactory
      End Get
    End Property

    Public Property SetAnnotations() As Boolean
      Get
        Return mblnSetAnnotations
      End Get
      Set(value As Boolean)
        mblnSetAnnotations = value
      End Set
    End Property

    ''' <summary>Determines whether or not the permissions will be set on the destination document using the permissions collection defined in the source document.</summary>
    ''' <remarks>This property can only be used on imports with providers that implement IUpdatePermissions.</remarks>
    Public Property SetPermissions As Boolean
      Get
        Return mblnSetPermissions
      End Get
      Set(value As Boolean)
        mblnSetPermissions = value
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

    Public Property VersionType As VersionTypeEnum
      Get
        Return menuVersionType
      End Get
      Set(value As VersionTypeEnum)
        menuVersionType = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpDocument As Document)
      Me.New(lpDocument, "", Core.FilingMode.UnFiled, Nothing, Nothing)
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpPathFactory As PathFactory)
      Me.New(lpDocument, "", lpPathFactory.FilingMode, lpPathFactory, Nothing)
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpFilingMode As FilingMode, ByVal lpPathFactory As PathFactory)
      Me.New(lpDocument, "", lpFilingMode, lpPathFactory, Nothing)
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpErrorMessage As String, ByVal lpFilingMode As FilingMode, ByVal lpPathFactory As PathFactory, ByVal lpWorker As BackgroundWorker)
      Me.New(lpDocument, "", lpFilingMode, lpPathFactory, Nothing, True)
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpErrorMessage As String, ByVal lpFilingMode As FilingMode, ByVal lpPathFactory As PathFactory, ByVal lpWorker As BackgroundWorker, ByVal lpSetAnnotations As Boolean)
      Me.New(lpDocument, lpErrorMessage, lpFilingMode, lpPathFactory, lpWorker, lpSetAnnotations, VersionTypeEnum.Unspecified)
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpErrorMessage As String, ByVal lpFilingMode As FilingMode, ByVal lpPathFactory As PathFactory, ByVal lpWorker As BackgroundWorker, ByVal lpSetAnnotations As Boolean, lpVersionType As VersionTypeEnum)
      MyBase.New(lpWorker)

      Document = lpDocument
      ErrorMessage = lpErrorMessage
      menuFilingMode = lpFilingMode
      mobjPathFactory = lpPathFactory
      mblnSetAnnotations = lpSetAnnotations
      menuVersionType = lpVersionType
    End Sub

#End Region

#Region "Private Methods"

    Protected Friend Overridable Function DebuggerIdentifier() As String
      Try
        Dim lobjReturnBuilder As New StringBuilder

        If Document IsNot Nothing Then
          lobjReturnBuilder.AppendFormat("{0}; ", Document.DebuggerIdentifier)
        End If

        lobjReturnBuilder.AppendFormat("FilingMode={0}; ", FilingMode.ToString)

        If PathFactory IsNot Nothing Then
          lobjReturnBuilder.AppendFormat("Path={0}; ", PathFactory.DebuggerIdentifier)
        End If

        If SetAnnotations Then
          lobjReturnBuilder.Append("Set Annotations")
        Else
          lobjReturnBuilder.Remove(lobjReturnBuilder.Length - 2, 2)
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
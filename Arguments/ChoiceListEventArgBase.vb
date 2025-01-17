'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports Documents.Core.ChoiceLists
Imports Documents.Utilities

#End Region

Namespace Arguments

  ''' <summary>
  ''' Contains all the parameters for the DocumentImported Event
  ''' </summary>
  ''' <remarks></remarks>
  Public MustInherit Class ChoiceListEventArgBase
    Inherits BackgroundEventArgs

#Region "Class Variables"

    Private ReadOnly mstrID As String
    Private mobjChoiceList As ChoiceList

#End Region

#Region "Public Properties"

    Public ReadOnly Property ID() As String
      Get
        Return mstrID
      End Get
    End Property

    Public ReadOnly Property ChoiceList() As ChoiceList
      Get
        Return mobjChoiceList
      End Get
    End Property

#End Region

#Region "Public Methods"

    Public Sub SetChoiceList(ByVal lpChoiceList As ChoiceList)
      Try
        mobjChoiceList = lpChoiceList
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpId As String)
      MyBase.New(Now)
      mstrID = lpId
    End Sub

    Public Sub New(ByVal lpChoiceList As ChoiceList)
      Me.New(lpChoiceList, lpChoiceList.Id, Now, Nothing)
    End Sub

    Public Sub New(ByVal lpChoiceList As ChoiceList, ByVal lpNewId As String)
      Me.New(lpChoiceList, lpNewId, Now, Nothing)
    End Sub

    Public Sub New(ByVal lpChoiceList As ChoiceList, ByVal lpTime As DateTime)
      Me.New(lpChoiceList, lpChoiceList.Id, lpTime, Nothing)
    End Sub

    Public Sub New(ByVal lpChoiceList As ChoiceList, ByVal lpNewId As String, ByVal lpTime As DateTime)
      Me.New(lpChoiceList, lpNewId, lpTime, Nothing)
    End Sub

    Public Sub New(ByVal lpExportChoiceListArgs As ExportChoiceListEventArgs, ByVal lpNewId As String)
      Me.New(lpExportChoiceListArgs, lpNewId, Now)
    End Sub

    Public Sub New(ByVal lpExportChoiceListArgs As ExportChoiceListEventArgs, ByVal lpNewId As String, ByVal lpTime As DateTime)
      Me.New(lpExportChoiceListArgs.ChoiceList, lpNewId, lpTime, lpExportChoiceListArgs.Worker)
    End Sub

    Public Sub New(ByVal lpChoiceList As ChoiceList, ByVal lpNewId As String, ByVal lpTime As DateTime, ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpTime, lpWorker)
      mstrID = lpNewId
      If lpChoiceList Is Nothing Then
        Throw New ArgumentException(Nothing, NameOf(lpChoiceList))
      End If
      mobjChoiceList = lpChoiceList
    End Sub

    Public Sub New(ByVal lpChoiceList As ChoiceList, ByVal lpNewId As String, ByVal lpTime As DateTime, ByRef lpErrorMessage As String, ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpTime, lpErrorMessage, lpWorker)
      mstrID = lpNewId
      If lpChoiceList Is Nothing Then
        Throw New ArgumentException(Nothing, NameOf(lpChoiceList))
      End If
      mobjChoiceList = lpChoiceList
    End Sub

#End Region

  End Class

End Namespace
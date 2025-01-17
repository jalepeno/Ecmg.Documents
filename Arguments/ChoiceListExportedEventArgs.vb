'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports Documents.Core.ChoiceLists

#End Region

Namespace Arguments

  ''' <summary>
  ''' Contains all the parameters for the DocumentImported Event
  ''' </summary>
  ''' <remarks></remarks>
  Public Class ChoiceListExportedEventArgs
    Inherits ChoiceListEventArgBase

#Region "Class Variables"

#End Region

#Region "Public Properties"

    'Public ReadOnly Property Name() As String
    '  Get
    '    Return MyBase.ID
    '  End Get
    'End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpChoiceList As ChoiceList)
      MyBase.New(lpChoiceList)
    End Sub

    Public Sub New(ByVal lpChoiceList As ChoiceList,
                   ByVal lpTime As DateTime)
      MyBase.New(lpChoiceList, lpChoiceList.Id, lpTime, Nothing)
    End Sub

    Public Sub New(ByVal lpExportChoiceListArgs As ExportChoiceListEventArgs)
      MyBase.New(lpExportChoiceListArgs.ChoiceList)
    End Sub

    Public Sub New(ByVal lpExportChoiceListArgs As ExportChoiceListEventArgs,
                   ByVal lpTime As DateTime)
      MyBase.New(lpExportChoiceListArgs.ChoiceList, lpTime)
    End Sub

    Public Sub New(ByVal lpChoiceList As ChoiceList,
                   ByVal lpNewId As String,
                   ByVal lpTime As DateTime,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpChoiceList, lpNewId, lpTime, lpWorker)
    End Sub

    Public Sub New(ByVal lpChoiceList As ChoiceList,
                   ByVal lpNewId As String,
                   ByVal lpTime As DateTime,
                   ByRef lpErrorMessage As String,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpChoiceList, lpNewId, lpTime, lpErrorMessage, lpWorker)
    End Sub

#End Region

  End Class

End Namespace
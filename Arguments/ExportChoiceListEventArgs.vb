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

  Public Class ExportChoiceListEventArgs
    Inherits ChoiceListEventArgBase

#Region "Constructors"

    Public Sub New(ByVal lpId As String)
      MyBase.New(lpId)
    End Sub

    Public Sub New(ByVal lpChoiceList As ChoiceList,
                   ByVal lpId As String)
      MyBase.New(lpChoiceList, lpId)
    End Sub

    Public Sub New(ByVal lpChoiceList As ChoiceList,
                   ByVal lpId As String,
                   ByVal lpTime As DateTime)
      MyBase.New(lpChoiceList, lpId, lpTime, Nothing)
    End Sub

    Public Sub New(ByVal lpExportChoiceListArgs As ExportChoiceListEventArgs,
                   ByVal lpId As String)
      MyBase.New(lpExportChoiceListArgs, lpId)
    End Sub

    Public Sub New(ByVal lpExportChoiceListArgs As ExportChoiceListEventArgs,
                   ByVal lpId As String,
                   ByVal lpTime As DateTime)
      MyBase.New(lpExportChoiceListArgs.ChoiceList, lpId, lpTime)
    End Sub

    Public Sub New(ByVal lpChoiceList As ChoiceList,
                   ByVal lpId As String,
                   ByVal lpTime As DateTime,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpChoiceList, lpId, lpTime, lpWorker)
    End Sub

    Public Sub New(ByVal lpChoiceList As ChoiceList,
                   ByVal lpId As String,
                   ByVal lpTime As DateTime,
                   ByRef lpErrorMessage As String,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpChoiceList, lpId, lpTime, lpErrorMessage, lpWorker)
    End Sub

#End Region

  End Class


End Namespace


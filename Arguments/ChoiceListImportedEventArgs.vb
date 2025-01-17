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
  Public Class ChoiceListImportedEventArgs
    Inherits ChoiceListEventArgBase

#Region "Class Variables"

#End Region

#Region "Public Properties"

    Public ReadOnly Property NewID() As String
      Get
        Return MyBase.ID
      End Get
    End Property

    Public ReadOnly Property NewChoiceList() As ChoiceList
      Get
        Return MyBase.ChoiceList
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpChoiceList As ChoiceList,
                   ByVal lpNewId As String)
      MyBase.New(lpChoiceList, lpNewId)
    End Sub

    Public Sub New(ByVal lpChoiceList As ChoiceList,
                   ByVal lpNewId As String,
                   ByVal lpTime As DateTime)
      MyBase.New(lpChoiceList, lpNewId, lpTime, Nothing)
    End Sub

    Public Sub New(ByVal lpImportChoiceListArgs As ImportChoiceListEventArgs,
                   ByVal lpNewId As String)
      MyBase.New(lpImportChoiceListArgs.ChoiceList, lpNewId)
    End Sub

    Public Sub New(ByVal lpImportChoiceListArgs As ImportChoiceListEventArgs,
                   ByVal lpNewId As String,
                   ByVal lpTime As DateTime)
      MyBase.New(lpImportChoiceListArgs.ChoiceList, lpNewId, lpTime)
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
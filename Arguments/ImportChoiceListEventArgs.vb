'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports System.ComponentModel
Imports Documents.Core.ChoiceLists

Namespace Arguments

  ''' <summary>
  ''' Contains all the parameters necessary for the ImportCVL method.
  ''' </summary>
  ''' <remarks>Used as a sole parameter for the ImportCVL method.</remarks>
  Public Class ImportChoiceListEventArgs
    Inherits BackgroundWorkerEventArgs

#Region "Class Variables"

    Private mobjChoiceList As ChoiceList
    Private mblnReplaceExisting As Boolean = False

#End Region

#Region "Public Properties"

    Public Property ChoiceList() As ChoiceList
      Get
        Return mobjChoiceList
      End Get
      Set(ByVal value As ChoiceList)
        mobjChoiceList = value
      End Set
    End Property

    Public Property ReplaceExisting() As Boolean
      Get
        Return mblnReplaceExisting
      End Get
      Set(ByVal value As Boolean)
        mblnReplaceExisting = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpChoiceList As ChoiceList)
      Me.New(lpChoiceList, False, "", Nothing)
    End Sub

    Public Sub New(ByVal lpChoiceList As ChoiceList,
      ByVal lpReplaceExisting As Boolean)

      Me.New(lpChoiceList, lpReplaceExisting, "", Nothing)

    End Sub

    Public Sub New(ByVal lpChoiceList As ChoiceList,
      ByVal lpReplaceExisting As Boolean,
      ByVal lpErrorMessage As String,
      ByVal lpWorker As BackgroundWorker)

      MyBase.New(lpWorker)

      ChoiceList = lpChoiceList
      ErrorMessage = lpErrorMessage
      ReplaceExisting = lpReplaceExisting

    End Sub

#End Region

  End Class

End Namespace
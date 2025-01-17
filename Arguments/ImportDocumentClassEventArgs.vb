'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports System.ComponentModel
Imports Documents.Core


Namespace Arguments
  ''' <summary>
  ''' Contains all the parameters necessary for the ImportDocumentClass method.
  ''' </summary>
  ''' <remarks>Used as a sole parameter for the ImportDocumentClass method.</remarks>
  Public Class ImportDocumentClassEventArgs
    Inherits BackgroundWorkerEventArgs

#Region "Class Variables"

    Private mobjDocumentClass As DocumentClass
    Private mblnReplaceExisting As Boolean

#End Region

#Region "Public Properties"

    Public Property DocumentClass() As DocumentClass
      Get
        Return mobjDocumentClass
      End Get
      Set(ByVal value As DocumentClass)
        mobjDocumentClass = value
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

    Public Sub New(ByVal lpDocumentClass As DocumentClass)
      Me.New(lpDocumentClass, False, "", Nothing)
    End Sub

    Public Sub New(ByVal lpDocumentClass As DocumentClass,
                   ByVal lpReplaceExisting As Boolean)

      Me.New(lpDocumentClass, lpReplaceExisting, "", Nothing)

    End Sub

    Public Sub New(ByVal lpDocumentClass As DocumentClass,
                   ByVal lpReplaceExisting As Boolean,
                   ByVal lpErrorMessage As String,
                   ByVal lpWorker As BackgroundWorker)

      MyBase.New(lpWorker)

      DocumentClass = lpDocumentClass
      ErrorMessage = lpErrorMessage
      ReplaceExisting = lpReplaceExisting

    End Sub

#End Region

  End Class
End Namespace
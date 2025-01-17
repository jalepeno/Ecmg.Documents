'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports System.ComponentModel

Namespace Arguments
  ''' <summary>
  ''' Base argument class for Cts operations that run on a background thread.
  ''' </summary>
  ''' <remarks></remarks>
  Public Class BackgroundWorkerEventArgs
    Inherits EventArgs

#Region "Class Variables"

    Private mobjWorker As BackgroundWorker
    Private mobjDoWorkEventArgs As DoWorkEventArgs
    Private mblnWorkInBackground As Boolean
    Private mstrErrorMessage As String = ""

#End Region

#Region "Public Properties"

    Public ReadOnly Property Worker() As BackgroundWorker
      Get
        Return mobjWorker
      End Get
    End Property

    Public ReadOnly Property DoWorkEventArgs() As DoWorkEventArgs
      Get
        Return mobjDoWorkEventArgs
      End Get
    End Property

    Public ReadOnly Property WorkInBackground() As Boolean
      Get
        Return mblnWorkInBackground
      End Get
    End Property

    Public Property ErrorMessage() As String
      Get
        Return mstrErrorMessage
      End Get
      Set(ByVal value As String)
        mstrErrorMessage = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      Me.New(Nothing)
      'mblnWorkInBackground = False
    End Sub

    Public Sub New(ByVal lpWorker As BackgroundWorker)

      If lpWorker IsNot Nothing Then
        mobjWorker = lpWorker
        mblnWorkInBackground = True
      End If

      Dim lobjArgument As Object = Nothing
      mobjDoWorkEventArgs = New ComponentModel.DoWorkEventArgs(lobjArgument)

    End Sub

#End Region

#Region "Public Methods"

    'Public Sub SetErrorMessage(ByVal lpErrorMessage As String)
    '  mstrErrorMessage = lpErrorMessage
    'End Sub

#End Region

  End Class
End Namespace
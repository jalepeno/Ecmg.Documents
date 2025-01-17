'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"


Imports System.ComponentModel
Imports Documents.Utilities

#End Region

Namespace Arguments

  ''' <summary>
  ''' Contains all the parameters for the base BackgroundEvent
  ''' </summary>
  ''' <remarks></remarks>
  Public Class BackgroundEventArgs
    Inherits EventArgs

#Region "Class Variables"

    Private mdatTime As DateTime
    Private mobjWorker As BackgroundWorker
    Private mobjDoWorkEventArgs As DoWorkEventArgs
    Private mblnWorkInBackground As Boolean
    Private mstrErrorMessage As String = String.Empty

#End Region

#Region "Public Properties"

    Public ReadOnly Property Time() As DateTime
      Get
        Return mdatTime
      End Get
    End Property

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

    Public ReadOnly Property ErrorMessage() As String
      Get
        Return mstrErrorMessage
      End Get
    End Property

#End Region

#Region "Public Methods"

    Public Sub SetErrorMessage(ByVal lpErrorMessage As String)
      Try
        mstrErrorMessage = lpErrorMessage
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpTime As DateTime)
      Me.New(lpTime, "", Nothing)
    End Sub

    Public Sub New(ByVal lpTime As DateTime, ByVal lpErrorMessage As String)
      Me.New(lpTime, lpErrorMessage, Nothing)
    End Sub

    Public Sub New(ByVal lpTime As DateTime, ByVal lpWorker As BackgroundWorker)
      Me.New(lpTime, "", lpWorker)
    End Sub

    Public Sub New(ByVal lpTime As DateTime, ByVal lpErrorMessage As String, ByVal lpWorker As BackgroundWorker)
      mdatTime = lpTime
      mobjWorker = lpWorker
      mstrErrorMessage = lpErrorMessage


      If lpWorker IsNot Nothing Then
        Dim lobjArgument As Object = Nothing
        mobjDoWorkEventArgs = New ComponentModel.DoWorkEventArgs(lobjArgument)
        mblnWorkInBackground = True
      End If

    End Sub

#End Region

  End Class

End Namespace
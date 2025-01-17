'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Arguments

  Public Class DocumentPluginEventArgs

#Region "Class Variables"

    Private mobjDocument As Document
    Private mdatTime As DateTime
    Private mstrErrorMessage As String = String.Empty

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

    Public ReadOnly Property Time() As DateTime
      Get
        Return mdatTime
      End Get
    End Property

    Public ReadOnly Property ErrorMessage() As String
      Get
        Return mstrErrorMessage
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      Try
        mdatTime = Now
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal document As Document)
      Try
        mobjDocument = document
        mdatTime = Now
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal document As Document, ByVal errorMessage As String)
      Try
        mobjDocument = document
        mstrErrorMessage = errorMessage
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal document As Document, ByVal time As DateTime, ByVal errorMessage As String)
      Try
        mobjDocument = document
        mdatTime = time
        mstrErrorMessage = errorMessage
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace
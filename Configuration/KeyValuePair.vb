'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Text
Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Configuration

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class KeyValuePair
    Implements INameValuePair

#Region "Class Variables"

    Private mstrKey As String = String.Empty
    Private mstrValue As String = String.Empty

#End Region

#Region "Constructors"
    Public Sub New()

    End Sub

    Public Sub New(ByVal lpKey As String, ByVal lpValue As String)
      Try
        mstrKey = lpKey
        mstrValue = lpValue
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Properties"

    Public Property Key() As String Implements INameValuePair.Name
      Get
        Try
          Return mstrKey
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As String)
        Try
          mstrKey = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property Value() As String Implements INameValuePair.Value
      Get
        Try
          Return mstrValue
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As String)
        Try
          mstrValue = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

#End Region

#Region "Protected Methods"

    Protected Friend Overridable Function DebuggerIdentifier() As String
      Dim lobjIdentifierBuilder As New StringBuilder
      Try
        If String.IsNullOrEmpty(Key) AndAlso String.IsNullOrEmpty(Value) Then
          Return "<Not Initialized>"
        ElseIf Not String.IsNullOrEmpty(Key) AndAlso String.IsNullOrEmpty(Value) Then
          Return String.Format("{0}:<No Value>", Key)
        ElseIf String.IsNullOrEmpty(Key) AndAlso Not String.IsNullOrEmpty(Value) Then
          Return String.Format("<No Key>:{0}", Value)
        ElseIf Key.Contains(":"c) OrElse Value.Contains(":"c) Then
          Return String.Format("{0}~{1}", Key, Value)
        Else
          Return String.Format("{0}:{1}", Key, Value)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace

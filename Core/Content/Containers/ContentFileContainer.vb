' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ContentFileContainer.vb
'  Description :  [type_description_here]
'  Created     :  3/11/2012 7:46:18 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.IO
Imports Documents.Utilities

#End Region

Namespace Core

  Public Class ContentFileContainer
    Inherits ContentContainerBase

#Region "Class Variables"

    Private mobjFileStream As FileStream = Nothing

#End Region

#Region "Public Properties"

    Public Overrides ReadOnly Property CanRead As Boolean
      Get
        Try
          If ContentStream IsNot Nothing AndAlso ContentStream.CanRead Then
            Return True
          Else
            Return False
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Property ContentStream As FileStream
      Get
        Return MyBase.FileContent
      End Get
      Set(ByVal value As FileStream)
        mobjFileStream = value
        MyBase.FileContent = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal lpFileName As String)
      MyBase.New()
      Try
        If File.Exists(lpFileName) = False Then
          Throw New Exceptions.InvalidPathException(
            String.Format("Unable to create content file container using path '{0}' the path is invalid.",
                          lpFileName), lpFileName)
        End If
        MyBase.FileName = lpFileName
        InitializeStream()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpContent As Content)
      MyBase.New(lpContent)
    End Sub

#End Region

#Region "Private Methods"

    Private Sub InitializeStream()
      Try
        mobjFileStream = New FileStream(Me.FileName, FileMode.Open)
        MyBase.FileContent = mobjFileStream
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

    '#Region "IDisposable Implementation"

    '    Protected Overrides Sub Dispose(disposing As Boolean)
    '      Try
    '        mobjFileStream.Close()
    '        mobjFileStream.Dispose()
    '        MyBase.Dispose(disposing)
    '      Catch ex As Exception
    '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '        ' Re-throw the exception to the caller
    '        Throw
    '      End Try
    '    End Sub

    '#End Region

  End Class

End Namespace

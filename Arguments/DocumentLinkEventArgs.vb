' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  Process.vb
'  Description :  [type_description_here]
'  Created     :  11/15/2012 3:41:55 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Arguments

  Public Class DocumentLinkEventArgs
    Inherits BackgroundWorkerEventArgs

#Region "Class Variables"

    Private mstrName As String = String.Empty
    Private mstrTargetUrl As String = String.Empty
    Private mobjProperties As IProperties = New ECMProperties
    Private mstrFileType As String = String.Empty
    Private mstrMimeType As String = String.Empty
    Private mstrId As String = String.Empty
    Private mobjTag As Object = Nothing
    Private mobjSourceDocument As Document = Nothing

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the name of the link.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Name As String
      Get
        Try
          Return mstrName
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '   Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrName = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '   Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the target url of the link.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TargetUrl As String
      Get
        Try
          Return mstrTargetUrl
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '   Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrTargetUrl = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '   Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the properties of the link.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Properties As IProperties
      Get
        Try
          Return mobjProperties
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '   Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As IProperties)
        Try
          mobjProperties = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '   Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the file type of the link.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FileType As String
      Get
        Try
          Return mstrFileType
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '   Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrFileType = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '   Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the mime type of the link.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property MimeType As String
      Get
        Try
          Return mstrMimeType
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '   Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrMimeType = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '   Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property ID As String
      Get
        Try
          Return mstrId
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '   Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrId = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '   Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property Tag As Object
      Get
        Try
          Return mobjTag
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '   Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Object)
        Try
          mobjTag = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '   Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property SourceDocument As Document
      Get
        Try
          Return mobjSourceDocument
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '   Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Document)
        Try
          mobjSourceDocument = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '   Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpName As String,
               ByVal lpTargetUrl As String,
               ByVal lpMimeType As String)
      Me.New(lpName, lpTargetUrl, lpMimeType, lpMimeType, Nothing)
    End Sub

    Public Sub New(ByVal lpName As String,
                   ByVal lpTargetUrl As String,
                   ByVal lpMimeType As String,
                   ByVal lpWorker As BackgroundWorker)
      Me.New(lpName, lpTargetUrl, lpMimeType, lpMimeType, lpWorker)
    End Sub

    Public Sub New(ByVal lpName As String,
                   ByVal lpTargetUrl As String,
                   ByVal lpFileType As String,
                   ByVal lpMimeType As String,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpWorker)
      Try
        mstrName = lpName
        mstrTargetUrl = lpTargetUrl
        mstrFileType = lpFileType
        mstrMimeType = lpMimeType
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '   Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace
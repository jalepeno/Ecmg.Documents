' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ContentTooLargeException.vb
'  Description :  Used in cases where the content is too large to handle.
'  Created     :  9/22/2011 4:58:20 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Exceptions

  ''' <summary>
  ''' To be thrown when we are unable to handle the size of the content file.
  ''' </summary>
  ''' <remarks></remarks>
  Public Class ContentTooLargeException
    Inherits CtsException

#Region "Class Variables"

    Private mobjContent As Content
    Private mstrContentSize As String = String.Empty

#End Region

#Region "Public Properties"

    Public ReadOnly Property Content As Content
      Get
        Return mobjContent
      End Get
    End Property

    Public ReadOnly Property ContentSize As String
      Get
        Try
          Return GetContentSize()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(content As Content)
      MyBase.New(GenerateDefaultMessage(content))
      Try
        mobjContent = content
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(contentSize As String)
      MyBase.New(GenerateDefaultMessage(contentSize))
      Try
        mstrContentSize = contentSize
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(content As Content, innerException As Exception)
      MyBase.New(GenerateDefaultMessage(content), innerException)
      Try
        mobjContent = content
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(contentSize As String, innerException As Exception)
      MyBase.New(GenerateDefaultMessage(contentSize), innerException)
      Try
        mstrContentSize = contentSize
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(content As Content, message As String)
      MyBase.New(message)
      Try
        mobjContent = content
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(contentSize As String, message As String)
      MyBase.New(message)
      Try
        mstrContentSize = contentSize
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(content As Content, message As String, innerException As Exception)
      MyBase.New(message, innerException)
      Try
        mobjContent = content
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(contentSize As String, message As String, innerException As Exception)
      MyBase.New(message, innerException)
      Try
        mstrContentSize = contentSize
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

    Private Function GetContentSize() As String
      Try

        If Content IsNot Nothing Then
          Return Content.FileSize.ToString
        Else
          Return mstrContentSize
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function GenerateDefaultMessage(content As Content) As String
      Try
        If content IsNot Nothing AndAlso (Not String.IsNullOrEmpty(content.FileName)) Then
          Return String.Format("The file '{0}' is too large ({1}) to handle in the current application.",
                               content.FileName, content.FileSize.ToString)
        Else
          Return "The file is too large to handle in the current application."
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function GenerateDefaultMessage(contentSize As String) As String
      Try
        Return String.Format("The file is too large ({0}) to handle in the current application.",
                                   contentSize)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
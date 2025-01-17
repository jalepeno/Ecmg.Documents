'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.IO
Imports System.Text
Imports Documents.Utilities

#End Region

Namespace Core

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Friend Class ZipFileStream

#Region "Public Properties"

    Public Property FileName As String
    Public Property DirectoryInArchive As String
    Public Property Stream As Stream

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpFileName As String, ByVal lpDirectoryInArchive As String, ByVal lpStream As Stream)
      Try
        FileName = lpFileName
        DirectoryInArchive = lpDirectoryInArchive
        Stream = lpStream
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

    Protected Friend Function DebuggerIdentifier() As String
      Try

        Dim lobjDebuggerBuilder As New StringBuilder


        If DirectoryInArchive IsNot Nothing OrElse DirectoryInArchive.Length > 0 Then
          lobjDebuggerBuilder.AppendFormat("{0}\", DirectoryInArchive)
        End If

        If Stream IsNot Nothing Then
          lobjDebuggerBuilder.AppendFormat("{0}: {1}", Me.FileName, New FileSize(Me.Stream.Length).ToString)
        Else
          lobjDebuggerBuilder.AppendFormat("{0}: Stream not initialized", Me.FileName)
        End If

        Return Helper.CleanPath(lobjDebuggerBuilder.ToString)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
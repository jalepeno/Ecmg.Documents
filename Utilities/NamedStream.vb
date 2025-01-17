' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  NamedStream.vb
'  Description :  [type_description_here]
'  Created     :  05/30/2012 3:24:18 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.IO
Imports System.Text

#End Region

Namespace Utilities

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class NamedStream
    Implements INamedStream

#Region "Class Variables"

    Private mstrFileName As String = String.Empty
    Private mobjStream As Stream = Nothing
#If CTSDOTNET = 1 Then
    Private mstrTempFilePath As String = String.Empty
#End If

#End Region

#Region "Public Properties"

    Public Property FileName As String Implements INamedStream.FileName
      Get
        Return mstrFileName
      End Get
      Set(ByVal value As String)
        mstrFileName = value
      End Set
    End Property

    Public Property Stream As Stream Implements INamedStream.Stream
      Get
        Return mobjStream
      End Get
      Set(ByVal value As Stream)
        mobjStream = value
      End Set
    End Property

#If CTSDOTNET = 1 Then

    Public ReadOnly Property TempFilePath As String Implements INamedStream.TempFilePath
      Get
        Try
          Return mstrTempFilePath
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End If

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpStream As Stream, ByVal lpFileName As String)
      Try
        Stream = lpStream
        FileName = lpFileName
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

#If CTSDOTNET = 1 Then

    Public Function WriteToTempFile() As String Implements INamedStream.WriteToTempFile
      Try

        ' We are going to create a dedicated guid based subfolder for this temp file
        Dim lstrTempFileFolder As String = Helper.CleanPath(String.Format("{0}\{1}",
          FileHelper.Instance.TempPath, Guid.NewGuid.ToString))
        If Not Directory.Exists(lstrTempFileFolder) Then
          Directory.CreateDirectory(lstrTempFileFolder)
        End If

        mstrTempFilePath = Helper.CleanPath(String.Format("{0}\{1}",
          lstrTempFileFolder, Me.FileName))
        mstrTempFilePath = Helper.CleanFile(mstrTempFilePath, "_")

        If File.Exists(mstrTempFilePath) Then
          If Helper.IsFileLocked(mstrTempFilePath) = True Then
            Throw New Exceptions.FileLockedException(mstrTempFilePath)
          End If
        End If
        Helper.WriteStreamToFile(Me.Stream, mstrTempFilePath, FileMode.Create)
        Return mstrTempFilePath
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub DeleteTempFile() Implements INamedStream.DeleteTempFile
      Try
        If File.Exists(mstrTempFilePath) Then
          If Helper.IsFileLocked(mstrTempFilePath) = True Then
            Throw New Exceptions.FileLockedException(mstrTempFilePath)
          Else
            Dim lstrTempFileFolder As String = Path.GetDirectoryName(mstrTempFilePath)
            File.Delete(mstrTempFilePath)
            ' Determine if the containing folder is empty and is a guid.
            ' If so we will delete the containing folder as well.
            Dim lobjDirectory As New DirectoryInfo(lstrTempFileFolder)
            If lobjDirectory.GetFiles.Count = 0 AndAlso lobjDirectory.GetDirectories.Count = 0 Then
              If Helper.IsGuid(lobjDirectory.Name, New Guid) Then
                lobjDirectory.Delete()
              End If
            End If
          End If
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overrides Function ToString() As String
      Try
        Return DebuggerIdentifier()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        Return MyBase.ToString()
      End Try
    End Function

#End If

#Region "Protected Methods"

    Protected Friend Overridable Function DebuggerIdentifier() As String

      Dim lobjIdentifierBuilder As New StringBuilder

      Try

        If Not String.IsNullOrEmpty(FileName) Then
          lobjIdentifierBuilder.AppendFormat("{0}: ", FileName)

        Else
          lobjIdentifierBuilder.Append("Name not set: ")
        End If

        If Stream IsNot Nothing Then
          Dim lobjSize As New FileSize(Stream.Length)
          If Stream.CanRead Then
            lobjIdentifierBuilder.AppendFormat("Readable ({0})", lobjSize.ToString)
          Else
            lobjIdentifierBuilder.AppendFormat("Can't Read ({0})", lobjSize.ToString)
          End If

        Else
          lobjIdentifierBuilder.Append("(Stream not initialized)")
        End If

        Return lobjIdentifierBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return lobjIdentifierBuilder.ToString
      End Try

    End Function

#End Region

#End Region

  End Class

End Namespace
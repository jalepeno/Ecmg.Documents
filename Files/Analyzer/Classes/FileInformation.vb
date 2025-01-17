'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  FileInformation.vb
'   Description :  [type_description_here]
'   Created     :  1/26/2015 2:05:30 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Security.Cryptography
Imports System.Text
Imports Documents.Utilities
Imports Newtonsoft.Json

#End Region

Namespace Files

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  <Serializable()> Public Class FileInformation
    Implements IFileInformation
    Implements IComparable

#Region "Class Variables"

    Private mstrExtension As String = String.Empty
    Private mstrFileType As String = String.Empty

    Private mstrBasicHash As String = String.Empty

#End Region

#Region "IFileInfo Implementation"

    <JsonProperty("ext")> Public Property Extension As String Implements IFileInformation.Extension
      Get
        Try
          Return mstrExtension
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrExtension = value
          mstrBasicHash = CreateInfoHash()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    <JsonProperty("desc")> Public Property FileType As String Implements IFileInformation.FileType
      Get
        Try
          Return mstrFileType
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrFileType = value
          mstrBasicHash = CreateInfoHash()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    <JsonProperty("mime")> Public Property MimeType As String Implements IFileInformation.MimeType

    <JsonProperty("pop")> Public Property Popularity As Rating Implements IFileInformation.Popularity

    <JsonProperty("ref")> Public Property RefUrl As String Implements IFileInformation.RefUrl

    <JsonProperty("rem")> Public Property Remarks As String Implements IFileInformation.Remarks

    <JsonIgnore()>
    Public ReadOnly Property BasicHash As String Implements IFileInformation.BasicHash
      Get
        Try
          Return mstrBasicHash
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(lpExtension As String,
                   lpFileType As String,
                   lpMimeType As String)
      Me.New(lpExtension, lpFileType, lpMimeType, Rating.C, String.Empty, String.Empty)
    End Sub

    Public Sub New(lpExtension As String,
                   lpFileType As String)
      Me.New(lpExtension, lpFileType, String.Empty, Rating.C, String.Empty, String.Empty)
    End Sub

    Public Sub New(lpExtension As String,
                   lpFileType As String,
                   lpMimeType As String,
                   lpPopularity As Rating,
                   lpRefUrl As String,
                   lpRemarks As String)
      Try
        Extension = lpExtension
        FileType = lpFileType
        MimeType = lpMimeType
        Popularity = lpPopularity
        RefUrl = lpRefUrl
        Remarks = lpRemarks
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
      Try

        Return String.Format("{0}: {1}", Extension, FileType).CompareTo(String.Format("{0}: {1}", obj.Extension, obj.FileType))

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Friend Methods"

    Protected Friend Overridable Function DebuggerIdentifier() As String
      Dim lobjIdentifierBuilder As New StringBuilder
      Try

        If Not String.IsNullOrEmpty(Extension) Then
          lobjIdentifierBuilder.AppendFormat("{0}: ", Extension)
        Else
          lobjIdentifierBuilder.Append("No Extension: ")
        End If

        lobjIdentifierBuilder.AppendFormat("{0} ", Popularity)

        If Not String.IsNullOrEmpty(MimeType) Then
          lobjIdentifierBuilder.AppendFormat("({0}) ", MimeType)
        End If

        If Not String.IsNullOrEmpty(FileType) Then
          lobjIdentifierBuilder.AppendFormat("- {0}", FileType)
        End If

        Return lobjIdentifierBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return lobjIdentifierBuilder.ToString
      End Try
    End Function

#End Region

#Region "Private Methods"

    Private Function CreateInfoHash() As String
      Try

        Dim lstrSourceData As String = String.Empty
        Dim lobjTempSource As Byte()
        Dim lobjTempHash As Byte()

        lstrSourceData = String.Format("{0}~{1}", Extension, FileType)

        ' Create a byte array from the source data.
        lobjTempSource = ASCIIEncoding.ASCII.GetBytes(lstrSourceData)

        ' Compute the hash
        lobjTempHash = New MD5CryptoServiceProvider().ComputeHash(lobjTempSource)

        Return Helper.ByteArrayToString(lobjTempHash)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
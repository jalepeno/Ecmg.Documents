'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  FileTestResult.vb
'   Description :  [type_description_here]
'   Created     :  1/28/2015 1:36:59 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Text
Imports Documents.Utilities

#End Region

Namespace Files

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class FileTestResult
    Implements IFileTestResult
    Implements IComparable

#Region "ITestResult Implementation"

    Public Property Extension As String Implements IFileTestResult.Extension

    Public Property FilePoints As String Implements IFileTestResult.FilePoints

    Public Property FileType As String Implements IFileTestResult.FileType

    Public Property MimeType As String Implements IFileTestResult.MimeType

    Public Property Percentage As Single Implements IFileTestResult.Percentage

    Public Property Points As Integer Implements IFileTestResult.Points

    Public Property Popularity As Rating Implements IFileTestResult.Popularity

#End Region

#Region "IComparable Implementation"

    Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
      Try
        If obj Is Nothing Then
          Throw New ArgumentNullException("obj")
        End If

        ' Flip the percentage to make the percentage sort descending
        Return Percentage.CompareTo(100 - DirectCast(obj, FileTestResult).Percentage)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Constuctors"

    Public Sub New()
    End Sub

    Public Sub New(lpFileInfo As IFileInformation)
      Try
        If lpFileInfo Is Nothing Then
          Throw New ArgumentNullException("lpFileInfo")
        End If

        FileType = lpFileInfo.FileType
        Extension = lpFileInfo.Extension
        MimeType = lpFileInfo.MimeType
        Popularity = lpFileInfo.Popularity
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(lpFileDefinition As IFileDefinition)
      Me.New(lpFileDefinition.FileInfo)
    End Sub

#End Region

#Region "Public Methods"

    Public Overloads Function ToString() As String Implements IFileTestResult.ToString
      Try
        Return DebuggerIdentifier()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Protected Methods"

    Protected Friend Overridable Function DebuggerIdentifier() As String
      Dim lobjIdentifierBuilder As New StringBuilder
      Try

        lobjIdentifierBuilder.AppendFormat("{0}-({1}): {2}% - {3}",
                                           Extension, Popularity, Percentage, FileType)

        Return lobjIdentifierBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return lobjIdentifierBuilder.ToString
      End Try
    End Function

#End Region

  End Class

End Namespace
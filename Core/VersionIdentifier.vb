' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IVersionIdentifier.vb
'  Description :  Used for identifying a version of a document or object.
'  Created     :  12/19/2011 8:23:32 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Core

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class VersionIdentifier
    Implements IVersionIdentifier

#Region "Class Variables"

    Private mobjMajorVersion As Object = 0
    Private mobjMinorVersion As Object = 0

#End Region

#Region "Public Properties"

    Public ReadOnly Property MajorVersion As Object Implements IVersionIdentifier.MajorVersion
      Get
        Return mobjMajorVersion
      End Get
    End Property

    Public ReadOnly Property MinorVersion As Object Implements IVersionIdentifier.MinorVersion
      Get
        Return mobjMinorVersion
      End Get
    End Property

    Public ReadOnly Property Version As Object Implements IVersionIdentifier.Version
      Get
        Try
          Return GetVersion()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpMajorVersion As Object, ByVal lpMinorVersion As Object)
      Try
        mobjMajorVersion = lpMajorVersion
        mobjMinorVersion = lpMinorVersion
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Protected Methods"

    Protected Friend Function DebuggerIdentifier() As String
      Try
        If Version IsNot Nothing Then
          Return Version.ToString
        Else
          Return "VersionIdentifier"
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Function GetVersion() As String
      Try
        If MajorVersion IsNot Nothing AndAlso MinorVersion IsNot Nothing Then
          Return String.Format("{0}.{1}", MajorVersion.ToString, MinorVersion.ToString)
        ElseIf MajorVersion IsNot Nothing AndAlso MinorVersion Is Nothing Then
          Return String.Format("{0}.0", MajorVersion.ToString)
        Else
          Return "0.0"
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

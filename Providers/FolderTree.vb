'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports System.Text
Imports Documents.Utilities

Namespace Providers

  Public Class FolderTree

#Region "Class Variables"

    Private mstrFolderPath As String
    Private mobjFolders As Collection
    Private mstrName As String

#End Region

#Region "Structures"

    Public Structure FolderInfo
      Implements IEquatable(Of FolderInfo)

#Region "Structure Variables"

      Dim Name As String
      Dim Path As String
      Dim Order As Integer
      Dim StartPosition As Integer
      Dim EndPosition As Integer

#End Region

#Region "Public Methods"

      Public Overloads Function Equals(ByVal other As FolderInfo) As Boolean Implements System.IEquatable(Of FolderInfo).Equals
        Try

          If other.Name = Me.Name AndAlso
            other.Path = Me.Path AndAlso
            other.Order = Me.Order AndAlso
            other.StartPosition = Me.StartPosition AndAlso
            other.EndPosition = Me.EndPosition Then
            Return True
          Else
            Return False
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Function

      Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Try
          If TypeOf obj Is FolderInfo Then
            Return Equals(CType(obj, FolderInfo))
          Else
            Return False
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Function

#End Region

#Region "Operators"

      Public Shared Operator =(ByVal lpFirstObject As FolderInfo, ByVal lpSecondObject As FolderInfo) As Boolean
        Try
          Return lpFirstObject.Equals(lpSecondObject)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Operator

      Public Shared Operator <>(ByVal lpFirstObject As FolderInfo, ByVal lpSecondObject As FolderInfo) As Boolean
        Try
          Return Not lpFirstObject.Equals(lpSecondObject)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Operator

#End Region

    End Structure

#End Region

#Region "Public Properties"

    Public ReadOnly Property FolderPath() As String
      Get
        FolderPath = mstrFolderPath
      End Get
    End Property

    Public ReadOnly Property Folders() As Collection
      Get
        Folders = mobjFolders
      End Get
    End Property

    Public ReadOnly Property Name As String
      Get
        Try
          Return mstrName
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpFolderPath As String)
      mstrFolderPath = lpFolderPath
      mobjFolders = New Collection()
      ParseFolderPath(lpFolderPath)
    End Sub

#End Region

#Region "Public Methods"

    Public Function RemoveLevel(lpTargetLevel As Integer) As FolderTree
      Try

        If lpTargetLevel > Folders.Count Then
          Throw New InvalidOperationException(String.Format("The requested target folder level '{0}' is higher than the current number of folder levels ({1}).", lpTargetLevel, Folders.Count))
        End If

        Dim lobjPathBuilder As New StringBuilder

        Dim lobjNewFolderPathCollection As Collection = Folders
        lobjNewFolderPathCollection.Remove(lpTargetLevel)

        For Each lobjFolderLevel As FolderInfo In lobjNewFolderPathCollection
          lobjPathBuilder.AppendFormat("/{0}", lobjFolderLevel.Name)
        Next

        Return New FolderTree(lobjPathBuilder.ToString())

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    Private Function ParseFolderPath(ByVal lpFolderPath As String) As Collection
      Try
        'Dim lstrFolderPath As String
        Dim lintFolderCount As Integer
        Dim lintFolderCounter As Integer
        'Dim lintCurrentPosition As Integer
        'Dim lintNextPosition As Integer
        'Dim lstrCurrentFolderName As String
        Dim lobjFolderInfo As New FolderInfo()


        Dim lstrFolders() As String
        Dim lstrPreviousPath As String = ""

        '    lstrFolderPath = lpFolderPath
        '    lintCurrentPosition = InStr(lstrFolderPath, "\")
        '    lintNextPosition = 1
        '    Do Until lintNextPosition = 0
        'lstrCurrentFolderName=mid$(lstrfolderpath
        '    Loop
        'parsefolerpath=

        lstrFolders = Split$(lpFolderPath, "/")

        lintFolderCount = UBound(lstrFolders)

        For lintFolderCounter = 0 To lintFolderCount
          If Len(lstrFolders(lintFolderCounter)) > 0 Then
            lobjFolderInfo.Name = lstrFolders(lintFolderCounter)
            lobjFolderInfo.Order = lintFolderCounter
            lobjFolderInfo.Path = lstrPreviousPath & "/" & lobjFolderInfo.Name
            lstrPreviousPath = lobjFolderInfo.Path
            Folders.Add(lobjFolderInfo, lobjFolderInfo.Order)
          End If
        Next
        ParseFolderPath = Folders
        mstrName = lstrFolders(lintFolderCount)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
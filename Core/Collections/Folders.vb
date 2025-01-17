'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Providers
Imports Documents.Utilities

Namespace Core
  ''' <summary>Collection of Folder objects.</summary>
  Public Class Folders
    Inherits CCollection(Of CFolder)

#Region "Class Variables"

    Dim mobjProvider As Providers.CProvider

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByRef lpProvider As Providers.CProvider)
      mobjProvider = lpProvider
    End Sub

#End Region

#Region "Public Properties"

    Public ReadOnly Property Provider() As Providers.CProvider
      Get
        Return mobjProvider
      End Get
    End Property

    Default Shadows Property Item(ByVal Path As String) As CFolder
      Get
        Try
          Dim lobjFolder As CFolder
          For lintCounter As Integer = 0 To MyBase.Count - 1
            lobjFolder = CType(MyBase.Item(lintCounter), CFolder)
            If lobjFolder.Path = Path Then
              Return lobjFolder
            End If
          Next
          Throw New Exception("There is no Item by the Path '" & Path & "'.")
          'Throw New InvalidArgumentException
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As CFolder)
        Try
          Dim lobjVersion As CFolder
          For lintCounter As Integer = 0 To MyBase.Count - 1
            lobjVersion = CType(MyBase.Item(lintCounter), CFolder)
            If lobjVersion.Path = Path Then
              MyBase.Item(lintCounter) = value
            End If
          Next
          Throw New Exception("There is no Item by the Path '" & Path & "'.")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Default Shadows Property Item(ByVal Index As Integer) As CFolder
      Get
        Return MyBase.Item(Index)
      End Get
      Set(ByVal value As CFolder)
        MyBase.Item(Index) = value
      End Set
    End Property

#End Region

#Region "Public Methods"

    'Public Overloads Sub Add(ByVal lpFolder As IFolder)
    '  MyBase.Add(lpFolder)
    'End Sub

    Public Shadows Sub Sort()
      Try
        MyBase.Sort(Comparer.Default)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub Sort(ByVal comparer As System.Collections.IComparer)
      Try
        MyBase.Sort(comparer)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace
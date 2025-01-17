' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  FolderClasses.vb
'  Description :  [type_description_here]
'  Created     :  3/5/2012 8:58:11 AM
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

  Public Class FolderClasses
    Inherits RepositoryObjectClasses

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(lpObjectClasses As RepositoryObjectClasses)
      Try

        For Each lobjObjectClass As RepositoryObjectClass In lpObjectClasses
          If TypeOf lobjObjectClass Is FolderClass Then
            MyBase.Add(lobjObjectClass)
          End If
        Next
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Shadows Sub Add(ByVal item As FolderClass)
      Try
        MyBase.Add(item)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub Add(ByVal items As FolderClasses)
      Try
        MyBase.Add(items)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Gets or sets the folder class based on the specified identifier.
    ''' </summary>
    ''' <param name="Name">Interpreted in order as Name, Label and ID.</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Default Shadows Property Item(ByVal Name As String) As FolderClass
      Get
        Return MyBase.Item(Name)
      End Get
      Set(ByVal value As FolderClass)
        MyBase.Item(Name) = value
      End Set
    End Property

    Default Public Shadows ReadOnly Property Item(ByVal name As String,
                                                  ByVal ignoreCase As Boolean) As FolderClass
      Get
        Return MyBase.Item(name, ignoreCase)
      End Get
    End Property

    Default Shadows Property Item(ByVal Index As Integer) As FolderClass
      Get
        Return MyBase.Item(Index)
      End Get
      Set(ByVal value As FolderClass)
        MyBase.Item(Index) = value
      End Set
    End Property

    Public Shadows Function Remove(ByVal Item As FolderClass) As Boolean
      Try
        ' Since sometimes we could have two different object references 
        ' to the same class, we will attempt to handle it here by the 
        ' assumption that if the class has the same name, then it is 
        ' the one we really care about.
        If Contains(Item.Name) Then
          Return MyBase.Remove(Me.Item(Item.Name))
        Else
          Return False
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
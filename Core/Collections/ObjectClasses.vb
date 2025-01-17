' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ObjectClasses.vb
'  Description :  [type_description_here]
'  Created     :  9/2/2015 2:00:11 AM
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

  Public Class ObjectClasses
    Inherits RepositoryObjectClasses

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(lpObjectClasses As RepositoryObjectClasses)
      Try

        For Each lobjObjectClass As RepositoryObjectClass In lpObjectClasses
          If TypeOf lobjObjectClass Is ObjectClass Then
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

    Public Shadows Sub Add(ByVal item As ObjectClass)
      Try
        MyBase.Add(item)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub Add(ByVal items As ObjectClasses)
      Try
        MyBase.Add(items)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Gets or sets the object class based on the specified identifier.
    ''' </summary>
    ''' <param name="Name">Interpreted in order as Name, Label and ID.</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Default Shadows Property Item(ByVal Name As String) As ObjectClass
      Get
        Return MyBase.Item(Name)
      End Get
      Set(ByVal value As ObjectClass)
        MyBase.Item(Name) = value
      End Set
    End Property

    Default Public Shadows ReadOnly Property Item(ByVal name As String,
                                                  ByVal ignoreCase As Boolean) As ObjectClass
      Get
        Return MyBase.Item(name, ignoreCase)
      End Get
    End Property

    Default Shadows Property Item(ByVal Index As Integer) As ObjectClass
      Get
        Return MyBase.Item(Index)
      End Get
      Set(ByVal value As ObjectClass)
        MyBase.Item(Index) = value
      End Set
    End Property

    Public Shadows Function Remove(ByVal Item As ObjectClass) As Boolean
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
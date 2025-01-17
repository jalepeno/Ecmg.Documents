'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Data
  ''' <summary>Defines the table, view or query to search.</summary>
  Public Class QueryTarget
    Implements IComparable

#Region "Class Variables"

    Dim mstrName As String = String.Empty
    Dim menuType As QueryTargetType
    Dim mstrQualifiedName As String = String.Empty

#End Region

#Region "Public Properties"

    Public Property Name() As String
      Get
        Try
          Return mstrName
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal Value As String)
        Try
          mstrName = Value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property QualifiedName() As String
      Get
        Try
          Return mstrQualifiedName
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal Value As String)
        Try
          mstrQualifiedName = Value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property Type() As QueryTargetType
      Get
        Try
          Return menuType
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal Value As QueryTargetType)
        Try
          menuType = Value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpName As String, ByVal lpType As QueryTargetType)
      Try
        Name = lpName
        QualifiedName = lpName
        Type = lpType
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpName As String, ByVal lpQualifiedName As String, ByVal lpType As QueryTargetType)
      Try
        Name = lpName
        QualifiedName = lpQualifiedName
        Type = lpType
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "IComparable Implementation"

    Public Function CompareTo(ByVal obj As Object) As Integer Implements IComparable.CompareTo

      If TypeOf obj Is QueryTarget Then
        Return Name.CompareTo(obj.Name)
      Else
        Throw New ArgumentException("Object is not a QueryTarget")
      End If

    End Function

#End Region

  End Class
End Namespace
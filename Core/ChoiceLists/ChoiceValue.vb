'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

Namespace Core.ChoiceLists

  ''' <summary>
  ''' An individual value in a ChoiceList
  ''' </summary>
  ''' <remarks></remarks>
  Public Class ChoiceValue
    Inherits ChoiceItem
    Implements IComparable

#Region "Class Variables"

    Private menuChoiceType As ChoiceType
    Private mobjValue As Object

#End Region

#Region "Public Properties"

    Public Property ChoiceType() As ChoiceType
      Get
        Return menuChoiceType
      End Get
      Set(ByVal value As ChoiceType)
        menuChoiceType = value
      End Set
    End Property

    Public Property Value() As Object
      Get
        Return mobjValue
      End Get
      Set(ByVal value As Object)
        mobjValue = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpValue As String)
      MyBase.New(lpValue)
      Try
        MyClass.Value = lpValue
        ChoiceType = ChoiceType.ChoiceString
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpValue As Integer)
      MyBase.New(lpValue)
      Try
        MyClass.Value = lpValue
        ChoiceType = ChoiceType.ChoiceInteger
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpValue As Nullable(Of Integer))
      MyBase.New(lpValue.GetValueOrDefault)
      Try
        MyClass.Value = lpValue.GetValueOrDefault
        ChoiceType = ChoiceType.ChoiceInteger
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

    Public Overrides Function ToString() As String
      Try
        Return Me.Value.ToString
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#Region "IComparable Implementation"

    Public Overloads Function CompareTo(ByVal obj As Object) As Integer
      Try
        If TypeOf obj Is ChoiceValue Then
          Return DisplayName.CompareTo(obj.DisplayName)
        Else
          Throw New ArgumentException("Object is not a ChoiceValue")
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
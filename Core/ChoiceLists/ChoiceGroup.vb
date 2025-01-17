'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

Namespace Core.ChoiceLists

  Public Class ChoiceGroup
    Inherits ChoiceItem
    Implements IComparable

#Region "Class Variables"
    Private mobjChoiceValues As New ChoiceValues
#End Region

#Region "Public Properties"

    Public Property ChoiceValues() As ChoiceValues
      Get
        Return mobjChoiceValues
      End Get
      Set(ByVal value As ChoiceValues)
        mobjChoiceValues = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal Name As String)
      MyBase.New(Name)
    End Sub

    Public Sub New(ByVal Name As String, ByVal Values As ChoiceValues)
      MyBase.New(Name)
      MyClass.ChoiceValues = Values
    End Sub

#End Region

#Region "IComparable Implementation"

    Public Overloads Function CompareTo(ByVal obj As Object) As Integer
      Try
        If TypeOf obj Is ChoiceGroup Then
          Return DisplayName.CompareTo(obj.DisplayName)
        Else
          Throw New ArgumentException("Object is not a ChoiceGroup")
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
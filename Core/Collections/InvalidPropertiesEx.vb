'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports Documents.Utilities

#End Region

Namespace Core

  Partial Public Class InvalidProperties
    Implements INamedItems
    Implements IProperties

#Region "Class Variables"

    Private mobjEnumerator As IEnumeratorConverter(Of IProperty)

#End Region

#Region "Private Properties"

    Private ReadOnly Property IPropertyEnumerator As IEnumeratorConverter(Of IProperty)
      Get
        Try
          'If mobjIPropertyEnumerator Is Nothing Then
          '  mobjIPropertyEnumerator = New IEnumeratorConverter(Of IProperty)(Me.ToArray, GetType(InvalidProperty), GetType(IProperty))
          'End If
          'Return mobjIPropertyEnumerator
          If mobjEnumerator Is Nothing OrElse mobjEnumerator.Count <> Me.Count Then
            mobjEnumerator = New IEnumeratorConverter(Of IProperty)(Me.ToArray, GetType(InvalidProperty), GetType(IProperty))
          End If
          Return mobjEnumerator
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

    Public Shadows Function Remove(ByVal lpProperty As InvalidProperty) As Boolean
      Try
        Return MyBase.Remove(lpProperty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    <XmlIgnore()>
    Property IItem(ByVal Name As String) As IProperty Implements IProperties.Item
      Get
        Try
          Return Item(Name)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As IProperty)
        Try
          Item(Name) = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Sub Replace(ByVal lpName As String, ByVal lpNewProperty As IProperty) Implements IProperties.Replace
      Try
        If PropertyExists(lpName) Then
          Remove(lpName)
        End If
        Add(lpNewProperty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#Region "INamedItems Implementation"

    Public Shadows Sub Add(ByVal item As Object) Implements INamedItems.Add
      Try
        If TypeOf (item) Is InvalidProperty Then
          Add(DirectCast(item, InvalidProperty))
        Else
          Throw New ArgumentException(
            String.Format("Invalid object type.  Item type of InvalidProperty expected instead of type {0}",
                          item.GetType.Name), "item")
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Property ItemByName(ByVal name As String) As Object Implements INamedItems.ItemByName
      Get
        Return Item(name)
      End Get
      Set(ByVal value As Object)
        Item(name) = value
      End Set
    End Property

    Public Shadows Function Remove(ByVal lpItemName As String) As Boolean Implements INamedItems.Remove
      Try
        ' Try to get the item
        Dim lobjItem As Object = Item(lpItemName)
        If lobjItem IsNot Nothing Then
          Return MyBase.Remove(Item(lpItemName))
        Else : Return False
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IProperties Implementation"

    Public Shadows Sub Add(ByVal lpProperties As IProperties) Implements IProperties.Add
      Try
        For Each lobjProperty As IProperty In lpProperties
          If Contains(lobjProperty.Name) = False Then
            Add(lobjProperty)
          End If
        Next
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub Add(ByVal item As IProperty) Implements System.Collections.Generic.ICollection(Of IProperty).Add
      Try
        Add(item)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Sub Clear() Implements System.Collections.Generic.ICollection(Of IProperty).Clear
      Try
        MyBase.Clear()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Function Contains(ByVal item As IProperty) As Boolean Implements System.Collections.Generic.ICollection(Of IProperty).Contains
      Try
        Return MyBase.Contains(item.Name)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Sub CopyTo(ByVal array() As IProperty, ByVal arrayIndex As Integer) Implements System.Collections.Generic.ICollection(Of IProperty).CopyTo
      Try
        MyBase.CopyTo(array, arrayIndex)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads ReadOnly Property Count As Integer Implements System.Collections.Generic.ICollection(Of IProperty).Count
      Get
        Try
          Return MyBase.Count
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Overloads ReadOnly Property IsReadOnly As Boolean Implements System.Collections.Generic.ICollection(Of IProperty).IsReadOnly
      Get
        Try
          Return MyBase.IsReadOnly
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Shadows Function Remove(ByVal item As IProperty) As Boolean Implements System.Collections.Generic.ICollection(Of IProperty).Remove
      Try
        Return Remove(item.Name)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of IProperty) Implements System.Collections.Generic.IEnumerable(Of IProperty).GetEnumerator
      Try
        Return IPropertyEnumerator.GetEnumerator()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Function PropertyExists(ByVal lpName As String, ByVal lpCaseSensitive As Boolean) As Boolean Implements IProperties.PropertyExists

      ' Does the property exist?
      Try
        Return PropertyExists(lpName, lpCaseSensitive, Nothing)
        'ApplicationLogging.WriteLogEntry("Enter Properties::PropertyExists", TraceEventType.Verbose)
        'For Each lobjProperty As ECMProperty In Me
        '  If (lpCaseSensitive = True) Then
        '    If lobjProperty.Name = lpName Then
        '      Return True
        '    End If
        '  Else
        '    If lobjProperty.Name.ToLower = lpName.ToLower Then
        '      Return True
        '    End If
        '  End If
        'Next
        'Return False
      Catch ex As Exception
        ApplicationLogging.LogException(ex, "Properties::PropertyExists")
        Return False
      Finally
        ApplicationLogging.WriteLogEntry("Exit Properties::PropertyExists", TraceEventType.Verbose)
      End Try
    End Function

#End Region

#Region "ICloneable Implementation"

    Public Function Clone() As Object Implements System.ICloneable.Clone

      Dim lobjProperties As New InvalidProperties

      Try
        For Each lobjProperty As InvalidProperty In Me
          lobjProperties.Add(lobjProperty.Clone)
        Next
        Return lobjProperties
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

#End Region

  End Class

End Namespace
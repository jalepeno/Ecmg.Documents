'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports System.ComponentModel
Imports Documents.Core
Imports Documents.Utilities

Namespace Providers
  ''' <summary>Collection of name/value pairs of properties for a CProvider object.</summary>
  Public Class ProviderProperties
    Inherits CCollection(Of ProviderProperty)

#Region "Class Variables"

    ' Class level variable for ProviderProperty
    ' Used to monitor change event for passing on
    Private WithEvents mobjProviderProperty As ProviderProperty

#End Region

#Region "Class Events"

    Public Event ItemValueChanged As ProviderPropertyValueChangedEventHandler
    Public Event ProviderPropertyChanged(sender As Object, e As PropertyChangedEventArgs)

#End Region

#Region "Overloads"

    ''' <summary>
    ''' Does a case insensitive check against all provider 
    ''' properties using the proeprty name to see if the 
    ''' specified property exists in the collection.
    ''' </summary>
    ''' <param name="name">The proeprty name to look for.</param>
    ''' <returns>True if a corresponding proeprty exists, otherwise false.</returns>
    ''' <remarks></remarks>
    Public Overloads Function Contains(ByVal name As String) As Boolean
      Try
        For Each lobjProviderProperty As ProviderProperty In Me
          If String.Equals(lobjProviderProperty.PropertyName, name, StringComparison.InvariantCultureIgnoreCase) Then
            Return True
          End If
        Next
        Return False
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Default Shadows Property Item(ByVal Name As String) As ProviderProperty
      Get
        Try
          'ApplicationLogging.WriteLogEntry(String.Format("Enter ProviderProperties::Get_Item('{0}' As String)", Name), TraceEventType.Verbose)
          For lintCounter As Integer = 0 To MyBase.Count - 1
            mobjProviderProperty = CType(MyBase.Item(lintCounter), ProviderProperty)
            If mobjProviderProperty.PropertyName = Name Then
              Return mobjProviderProperty
            End If
          Next
          'Throw New Exception(String.Format("There is no Item by the Name '{0}'.", Name))
          Return Nothing
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("ProviderProperties::Get_Item('{0}' As String)", Name), Name)
          Return Nothing
        Finally
          'ApplicationLogging.WriteLogEntry(String.Format("Exit ProviderProperties::Get_Item('{0}' As String)", Name), TraceEventType.Verbose)
        End Try
      End Get
      Set(ByVal value As ProviderProperty)
        Try
          'ApplicationLogging.WriteLogEntry(String.Format("Enter ProviderProperties::Set_Item('{0}' As String)", Name), TraceEventType.Verbose)
          Dim lobjProviderProperty As ProviderProperty
          For lintCounter As Integer = 0 To MyBase.Count - 1
            lobjProviderProperty = CType(MyBase.Item(lintCounter), ProviderProperty)
            If lobjProviderProperty.PropertyName = Name Then
              MyBase.Item(lintCounter) = value
            End If
          Next
          Throw New Exception("There is no Item by the Name '" & Name & "'.")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("ProviderProperties::Set_Item('{0}' As String)", Name), Name)
        Finally
          'ApplicationLogging.WriteLogEntry(String.Format("Exit ProviderProperties::Set_Item('{0}' As String)", Name), TraceEventType.Verbose)
        End Try
      End Set
    End Property

    Default Shadows Property Item(ByVal Index As Integer) As ProviderProperty
      Get
        Return MyBase.Item(Index)
      End Get
      Set(ByVal value As ProviderProperty)
        MyBase.Item(Index) = value
      End Set
    End Property

    Public Overloads Function Remove(ByVal lpPropertyName As String) As Boolean
      Try
        Dim lobjProviderProperty As ProviderProperty
        For lintCounter As Integer = 0 To MyBase.Count - 1
          lobjProviderProperty = CType(MyBase.Item(lintCounter), ProviderProperty)
          If String.Equals(lobjProviderProperty.PropertyName, lpPropertyName, StringComparison.InvariantCultureIgnoreCase) Then
            Remove(lobjProviderProperty)
            Return True
          End If
        Next
        Return False
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    Private Sub mobjProviderProperty_ProviderPropertyValueChanged(ByVal sender As Object, ByRef e As Arguments.ProviderPropertyValueChangedEventArgs) Handles mobjProviderProperty.ProviderPropertyValueChanged
      Try
        ' Raise the event to the next level
        RaiseEvent ItemValueChanged(Me, e)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Sub ProviderProperties_ItemPropertyChanged(sender As Object, e As ComponentModel.PropertyChangedEventArgs) Handles Me.ItemPropertyChanged
      Try
        RaiseEvent ProviderPropertyChanged(sender, e)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Nested Classes"

    Public Class ProviderPropertySequenceComparer
      Implements System.Collections.Generic.IComparer(Of ProviderProperty)
      Implements System.Collections.IComparer

      Public Function Compare(ByVal x As ProviderProperty, ByVal y As ProviderProperty) As Integer _
        Implements IComparer(Of ProviderProperty).Compare

        Try
          ' Two null objects are equal
          If (x Is Nothing) And (y Is Nothing) Then Return 0

          ' Any non-null object is greater than a null object
          If (x Is Nothing) Then Return 1
          If (y Is Nothing) Then Return -1

          If x.SequenceNumber < y.SequenceNumber Then
            Return -1
          Else
            Return 1
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try

      End Function

      Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare

        Try
          Return Compare(CType(x, ProviderProperty), CType(y, ProviderProperty))
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try

      End Function
    End Class

#End Region

  End Class

End Namespace
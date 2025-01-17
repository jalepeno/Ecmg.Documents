' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  Parameters.vb
'  Description :  Used to hold a collection of Parameters.
'  Created     :  11/17/2011 4:34:58 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports Documents.SerializationUtilities
Imports Documents.Utilities

#End Region

Namespace Core

  <Serializable()>
  Public Class Parameters
    Inherits CCollection(Of Parameter)
    Implements IParameters
    Implements IDisposable

#Region "Class Variables"

    Private mobjEnumerator As IEnumeratorConverter(Of IParameter)

#End Region

#Region "Public Events"

    Public Shadows Event ItemPropertyChanged(ByVal sender As Object, ByVal e As PropertyChangedEventArgs) Implements IParameters.ItemPropertyChanged

#End Region

#Region "Public Properties"

    Public Shadows Function GetItemByIndex(ByVal index As Integer) As IParameter Implements IParameters.GetItemByIndex
      Try
        Return MyBase.GetItemByIndex(index)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shadows Sub SetItemByIndex(ByVal index As Integer, ByVal value As IParameter) Implements IParameters.SetItemByIndex
      Try
        MyBase.SetItemByIndex(index, value)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Function GetItemByName(ByVal name As String) As IParameter Implements IParameters.GetItemByName
      Try
        Return MyBase.GetItemByName(name)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shadows Sub SetItemByName(ByVal name As String, ByVal value As IParameter) Implements IParameters.SetItemByName
      Try
        MyBase.SetItemByName(name, value)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Property Item(name As String) As IParameter Implements IParameters.Item
      Get
        Try
          Return MyBase.Item(name)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As IParameter)
        Try
          MyBase.Item(name) = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    'Default Overridable Shadows Property Item(ByVal name As String) As Parameter
    '  Get
    '    Try
    '      For Each lobjItem As IParameter In Me.Items
    '        If String.Compare(lobjItem.Name, name, True) = 0 Then
    '          Return lobjItem
    '        End If
    '      Next
    '      Return Nothing
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Get
    '  Set(value As Parameter)
    '    Try
    '      Dim lobjItem As Parameter = Item(value.Name)

    '      If lobjItem IsNot Nothing Then
    '        lobjItem.Value = value.Value
    '      End If

    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Set
    'End Property

#End Region

#Region "Private Properties"

    Protected ReadOnly Property IsDisposed() As Boolean
      Get
        Return disposedValue
      End Get
    End Property

    Private ReadOnly Property IPropertyEnumerator As IEnumeratorConverter(Of IParameter)
      Get
        Try
          If mobjEnumerator Is Nothing OrElse mobjEnumerator.Count <> Me.Count Then
            mobjEnumerator = New IEnumeratorConverter(Of IParameter)(Me.ToArray, GetType(Parameter), GetType(IParameter))
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

#Region "Public Methods"

    Public Overloads Function Clone() As Object Implements IParameters.Clone

      Dim lobjParameters As New Parameters

      Try
        For Each lobjParameter As Parameter In Me
          lobjParameters.Add(lobjParameter.Clone)
        Next
        Return lobjParameters
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Public Overrides Function Contains(ByVal lpName As String) As Boolean Implements IParameters.Contains
      Try
        Return MyBase.Contains(lpName)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Function GetValue(lpName As String) As Object Implements IParameters.GetValue
      Try
        If Contains(lpName) Then
          Dim lobjParameter As IParameter = Item(lpName)
          If lobjParameter Is Nothing OrElse lobjParameter.HasValue = False Then
            Return Nothing
          End If
          If lobjParameter.Cardinality = Cardinality.ecmSingleValued Then
            Return lobjParameter.Value
          Else
            Return lobjParameter.Values
          End If
        Else
          Return Nothing
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function Deserialize(ByVal lpFilePath As String, Optional ByRef lpErrorMessage As String = "") As Object
      Try
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
        Dim lobjParameters As Parameters = Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType)
        Return lobjParameters
      Catch ex As Exception
        'ApplicationLogging.LogException(ex, String.Format("{0}::Deserialize('{1}', '{2}')", Me.GetType.Name, lpFilePath, lpErrorMessage))
        lpErrorMessage = Helper.FormatCallStack(ex)
        Return Nothing
      End Try
    End Function

    Public Function DeSerialize(ByVal lpXML As System.Xml.XmlDocument) As Object
      Try
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
        Return Serializer.Deserialize.XmlString(lpXML.OuterXml, Me.GetType)
      Catch ex As Exception
        'ApplicationLogging.LogException(ex, String.Format("{0}::Deserialize(lpXML)", Me.GetType.Name))
        'Helper.DumpException(ex)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function Serialize() As System.Xml.XmlDocument
      Try
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
        Return Serializer.Serialize.Xml(Me)
      Catch ex As Exception
        'ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Sub Serialize(ByVal lpFilePath As String)
      Try
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
        Serializer.Serialize.XmlFile(Me, lpFilePath)
      Catch ex As Exception
        'ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "ICollection(Of IParameter) Implementation"

    Public Overloads Sub Add(item As IParameter) Implements System.Collections.Generic.ICollection(Of IParameter).Add
      Try
        Add(CType(item, Parameter))
        'If String.IsNullOrEmpty(item.Name()) Then
        '  Exit Sub
        'End If
        'If TypeOf item Is SingletonParameter OrElse TypeOf item Is MultiValueParameter OrElse String.IsNullOrEmpty(item.XsiType) Then
        '  If Contains(item.Name) = False Then
        '    MyBase.Add(CType(item, Parameter))
        '  End If
        'Else
        '  Dim lobjEcmProperty As Parameter = ParameterFactory.Create(item)
        '  If Contains(lobjEcmProperty.Name) = False Then
        '    MyBase.Add(lobjEcmProperty)
        '  End If
        'End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Sub Add(item As Parameter)
      Try
        ' MyBase.Add(item)
        If String.IsNullOrEmpty(item.Name()) Then
          Exit Sub
        End If
        If TypeOf item Is SingletonParameter OrElse TypeOf item Is MultiValueParameter Then
          If Contains(item.Name) = False Then
            MyBase.Add(item)
          End If
        Else
          Dim lobjEcmProperty As Parameter = ParameterFactory.Create(item)
          If Contains(lobjEcmProperty.Name) = False Then
            MyBase.Add(lobjEcmProperty)
          End If
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Sub AddRange(items As IParameters) Implements IParameters.AddRange
      Try
        'MyBase.AddRange(items)
        For Each lobjParameter As IParameter In items
          If GetItemByName(lobjParameter.Name) Is Nothing Then
            Add(lobjParameter)
          End If
        Next
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Sub Clear() Implements System.Collections.Generic.ICollection(Of IParameter).Clear
      Try
        MyBase.Clear()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Function Contains(item As IParameter) As Boolean Implements System.Collections.Generic.ICollection(Of IParameter).Contains
      Try
        Return Contains(CType(item, Parameter))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function Contains(item As Parameter) As Boolean
      Try
        Return MyBase.Contains(item)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Sub CopyTo(array() As IParameter, arrayIndex As Integer) Implements System.Collections.Generic.ICollection(Of IParameter).CopyTo
      Try
        MyBase.CopyTo(array, arrayIndex)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads ReadOnly Property Count As Integer Implements System.Collections.Generic.ICollection(Of IParameter).Count
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

    Public Overloads ReadOnly Property IsReadOnly As Boolean Implements System.Collections.Generic.ICollection(Of IParameter).IsReadOnly
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

    Public Overloads Function Remove(item As IParameter) As Boolean Implements System.Collections.Generic.ICollection(Of IParameter).Remove
      Try
        Return Remove(CType(item, Parameter))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function Remove(item As Parameter) As Boolean
      Try
        Return MyBase.Remove(item)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of IParameter) Implements System.Collections.Generic.IEnumerable(Of IParameter).GetEnumerator
      Try
        Return IPropertyEnumerator.GetEnumerator
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IDisposable Support"

    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
      If Not Me.disposedValue Then
        If disposing Then
          ' DISPOSETODO: dispose managed state (managed objects).
          MyBase.Dispose(disposing)
        End If

        ' DISPOSETODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
        ' DISPOSETODO: set large fields to null.
      End If
      Me.disposedValue = True
    End Sub

#End Region

  End Class

End Namespace
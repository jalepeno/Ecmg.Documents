'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Text
Imports System.Xml.Serialization
Imports Documents.Utilities

#End Region

Namespace Core
  ''' <summary>Collection of ObjectProperty objects.</summary>
  <Serializable()>
  Partial Public Class ECMProperties
    Implements ICloneable
    Implements INamedItems
    Implements IProperties
    Implements IDisposable
    Implements IEnumerable(Of IProperty)

#Region "Class Variables"

    'Private mintCurrentIndex As Integer = -1
    'Private mobjCurentItem As IProperty = Nothing
    Private mobjEnumerator As IEnumeratorConverter(Of IProperty)

#End Region

#Region "Private Properties"

    Protected ReadOnly Property IsDisposed() As Boolean
      Get
        Return disposedValue
      End Get
    End Property

    Private ReadOnly Property IPropertyEnumerator As IEnumeratorConverter(Of IProperty)
      Get
        Try
          If mobjEnumerator Is Nothing OrElse mobjEnumerator.Count <> Me.Count Then
            mobjEnumerator = New IEnumeratorConverter(Of IProperty)(Me.ToArray, GetType(ECMProperty), GetType(IProperty))
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

#Region "Constructors"

    'Public Sub New()
    '  Try
    '    mobjIPropertyEnumerator = New IEnumeratorConverter(Of IProperty)(Me, GetType(ECMProperty), GetType(IProperty))
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

#End Region

#Region "Public Methods"

    Public Shadows Sub Add(ByVal item As IProperty) Implements System.Collections.Generic.ICollection(Of IProperty).Add
      Try
        If TypeOf item Is ECMProperty Then
          If PropertyExists(item.Name) = False Then
            MyBase.Add(DirectCast(item, ECMProperty))
          End If
        Else
          ApplicationLogging.WriteLogEntry(
            String.Format("Unable to add property '{0}' to collection, it must inherit from ECMProperty.", item.Name),
            TraceEventType.Warning, 62914)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ' ''' <summary>
    ' ''' Gets or sets the property using the property name
    ' ''' </summary>
    ' ''' <param name="Name">The name of the property</param>
    ' ''' <value></value>
    ' ''' <returns>An ECMProperty reference if the property name
    ' ''' is found, otherwise returns null.</returns>
    ' ''' <remarks></remarks>
    '<XmlElement("Property", GetType(ECMProperty))> _
    '<DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    'Default Shadows Property Item(ByVal Name As String) As ECMProperty
    '  Get
    '    Dim lobjProperty As ECMProperty
    '    ' First look for the class by Name
    '    For lintCounter As Integer = 0 To MyBase.Count - 1
    '      lobjProperty = CType(MyBase.Item(lintCounter), ECMProperty)
    '      If String.Compare(lobjProperty.Name, Name, True) = 0 Then
    '        Return lobjProperty
    '      End If
    '    Next

    '    ' Next look for the class by PackedName
    '    For lintCounter As Integer = 0 To MyBase.Count - 1
    '      lobjProperty = CType(MyBase.Item(lintCounter), ECMProperty)
    '      If String.Compare(lobjProperty.PackedName, Name, True) = 0 Then
    '        Return lobjProperty
    '      End If
    '    Next

    '    'Throw New Exception("There is no Item by the Name '" & Name & "'.")
    '    'Throw New InvalidArgumentException

    '    ' If the property can't be found, just return Nothing
    '    Return Nothing

    '  End Get
    '  Set(ByVal value As ECMProperty)
    '    Dim lobjProperty As ECMProperty
    '    For lintCounter As Integer = 0 To MyBase.Count - 1
    '      lobjProperty = CType(MyBase.Item(lintCounter), ECMProperty)
    '      If String.Compare(lobjProperty.Name, Name, True) = 0 Then
    '        MyBase.Item(lintCounter) = value
    '      End If
    '    Next
    '    Throw New Exception("There is no Item by the Name '" & Name & "'.")
    '  End Set
    'End Property

    <XmlElement("Property", GetType(ECMProperty))>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Default Shadows Property Item(ByVal Index As Integer) As ECMProperty
      Get
        Return MyBase.Item(Index)
      End Get
      Set(ByVal value As ECMProperty)
        MyBase.Item(Index) = value
      End Set
    End Property

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

    'Public Shadows Sub AddRange(ByVal items As ECMProperties)
    '  Try
    '    For Each lobjItem As ECMProperty In items
    '      If Contains(lobjItem.Name) = False Then
    '        MyBase.Add(lobjItem)
    '      End If
    '    Next
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    Public Shadows Sub Add(ByVal item As ECMProperty)
      Try
        If String.IsNullOrEmpty(item.Name()) Then
          Exit Sub
        End If
        If TypeOf item Is SingletonProperty OrElse TypeOf item Is MultiValueProperty OrElse String.IsNullOrEmpty(item.XsiType) Then
          If Contains(item.Name) = False Then
            MyBase.Add(item)
          End If
        Else
          Dim lobjEcmProperty As ECMProperty = PropertyFactory.Create(item)
          If Contains(lobjEcmProperty.Name) = False Then
            MyBase.Add(lobjEcmProperty)
          Else
            ' Update the existing property
            Dim lobjExistingProperty As ECMProperty = ItemByName(item.Name)
            lobjExistingProperty.Value = item.Value
          End If
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub Add(ByVal lpProperties As IProperties) Implements IProperties.Add
      Try
        If TypeOf lpProperties Is ECMProperties Then
          Add(CType(lpProperties, ECMProperties))
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub Add(ByVal Properties As ECMProperties)
      Try
        For Each lobjEcmProperty As ECMProperty In Properties
          If Me.PropertyExists(lobjEcmProperty.Name) Then
            Select Case lobjEcmProperty.Cardinality
              Case Cardinality.ecmSingleValued
                If lobjEcmProperty.Value IsNot Nothing Then
                  Dim lobjOriginalProperty As ECMProperty = Me(lobjEcmProperty.Name)
                  If lobjEcmProperty.Cardinality = lobjOriginalProperty.Cardinality Then
                    If lobjOriginalProperty.Value Is Nothing OrElse lobjOriginalProperty.Value = String.Empty Then
                      lobjOriginalProperty.Value = lobjEcmProperty.Value
                    Else
                      ApplicationLogging.WriteLogEntry(String.Format("The property '{0}' is already defined in the collection, the value '{1}' was not added.",
                                                                  lobjEcmProperty.Name, lobjEcmProperty.Value.ToString), TraceEventType.Warning)
                    End If
                  End If
                Else
                  ApplicationLogging.WriteLogEntry(String.Format("The property '{0}' is already defined in the collection, the value was not added.",
                                                              lobjEcmProperty.Name), TraceEventType.Warning)
                End If
              Case Cardinality.ecmMultiValued
                ApplicationLogging.WriteLogEntry(String.Format("The property '{0}' is already defined in the collection, the values were not added.",
                                                            lobjEcmProperty.Name), TraceEventType.Warning)
            End Select
          Else
            Add(lobjEcmProperty)
          End If
        Next
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overrides Function Contains(ByVal lpName As String) As Boolean
      Try
        Return PropertyExists(lpName)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function Delete(ByVal lpName As String) As Boolean
      Try
        'Return Remove(ItemByName(lpName))
        Return Remove(lpName)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shadows Function Remove(ByVal lpProperty As ECMProperty) As Boolean
      Try
        Return MyBase.Remove(lpProperty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

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

    'Friend Function ToHashableString() As String
    '  Try

    '    Dim lobjStringBuilder As New StringBuilder

    '    For Each lobjProperty As ECMProperty In Me
    '      lobjStringBuilder.Append(lobjProperty.ToString)
    '    Next
    '    Return lobjStringBuilder.ToString

    '  Catch ex As Exception
    '     ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    Public Function ToXmlString() As String
      Try
        Return SerializationUtilities.Serializer.Serialize.XmlString(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ListProperties() As String
      Try

        Dim lobjStringBuilder As New StringBuilder

        For Each lobjProperty As ECMProperty In Me
          lobjStringBuilder.Append(String.Format("{0}, ", lobjProperty.Name))
        Next

        Return lobjStringBuilder.ToString.TrimEnd(" ").TrimEnd(",")

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    ''' <summary>
    ''' Used to syncronize the IProperty enumerator collection with the current collection
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ECMProperties_CollectionChanged(sender As Object, e As System.Collections.Specialized.NotifyCollectionChangedEventArgs) Handles Me.CollectionChanged
      Try
        Select Case e.Action
          Case Specialized.NotifyCollectionChangedAction.Add
            For Each lobjItem As Object In e.NewItems
              IPropertyEnumerator.Add(lobjItem)
            Next
          Case Specialized.NotifyCollectionChangedAction.Remove
            For Each lobjItem As Object In e.OldItems
              IPropertyEnumerator.Remove(lobjItem)
            Next
          Case Specialized.NotifyCollectionChangedAction.Reset
            IPropertyEnumerator.Reset()
        End Select
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "ICloneable Implementation"

    Public Function Clone() As Object Implements System.ICloneable.Clone

      Dim lobjProperties As New ECMProperties

      Try
        For Each lobjProperty As ECMProperty In Me
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

#Region "INamedItems Implementation"

    Public Shadows Sub Add(ByVal item As Object) Implements INamedItems.Add
      Try
        If TypeOf (item) Is ECMProperty Then
          Add(DirectCast(item, ECMProperty))
        Else
          Throw New ArgumentException(
            String.Format("Invalid object type.  Item type of ECMProperty expected instead of type {0}",
                          item.GetType.Name), "item")
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
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
        If Contains(lpItemName) Then
          Return MyBase.Remove(Item(lpItemName))
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IProperties Implementation"

    Public Shadows Sub Clear() Implements System.Collections.Generic.ICollection(Of IProperty).Clear
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
        Return Contains(item.Name)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shadows Sub CopyTo(ByVal array() As IProperty, ByVal arrayIndex As Integer) Implements System.Collections.Generic.ICollection(Of IProperty).CopyTo
      Try
        MyBase.CopyTo(array, arrayIndex)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows ReadOnly Property Count As Integer Implements System.Collections.Generic.ICollection(Of IProperty).Count
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

    Public Overrides ReadOnly Property IsReadOnly As Boolean Implements System.Collections.Generic.ICollection(Of IProperty).IsReadOnly
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
        Return IPropertyEnumerator.GetEnumerator
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Function IPropertyExists(ByVal lpName As String, ByVal lpCaseSensitive As Boolean) As Boolean Implements IProperties.PropertyExists

      ' Does the property exist?
      Try
        Return PropertyExists(lpName, lpCaseSensitive, Nothing)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, "Properties::PropertyExists")
        Return False
      Finally
        ApplicationLogging.WriteLogEntry("Exit Properties::PropertyExists", TraceEventType.Verbose)
      End Try
    End Function

#End Region

#Region "IEnumerator(Of IProperty) Implementation"

    'Public ReadOnly Property CurrentObject As Object Implements System.Collections.IEnumerator.Current
    '  Get
    '    Try
    '      Return Current
    '      'Return DirectCast(Me, IEnumerator).Current
    '      'MyBase.
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Get
    'End Property

    'Public ReadOnly Property Current As IProperty Implements System.Collections.Generic.IEnumerator(Of IProperty).Current
    '  Get
    '    Try
    '      If mobjCurentItem Is Nothing Then
    '        Throw New InvalidOperationException
    '      End If
    '      Return mobjCurentItem
    '      'Return DirectCast(Me, IEnumerator).Current
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Get
    'End Property

    'Public Function MoveNext() As Boolean Implements System.Collections.IEnumerator.MoveNext
    '  Try
    '    mintCurrentIndex += 1
    '    If mintCurrentIndex >= Me.Count Then
    '      ' Avoids going beyond the end of the collection.
    '      Return False
    '    Else
    '      'Set current item to next item in collection.
    '      mobjCurentItem = Item(mintCurrentIndex)
    '    End If
    '    'DirectCast(mybase, IEnumerator).MoveNext()

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    'Public Sub Reset() Implements System.Collections.IEnumerator.Reset
    '  Try
    '    mintCurrentIndex = -1
    '    'DirectCast(Items, IEnumerator).Reset()
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

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
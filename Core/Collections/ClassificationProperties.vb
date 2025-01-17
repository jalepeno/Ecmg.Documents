'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Reflection
Imports System.Xml.Serialization
Imports Documents.Utilities

#End Region

Namespace Core

  '<DebuggerDisplay("Classification Property Collection")> _
  ''' <summary>Collection of ClassificationProperty objects.</summary>
  Public Class ClassificationProperties
    Inherits CCollection(Of ClassificationProperty)
    Implements INamedItems
    Implements IProperties

#Region "Class Variables"

    Private mobjEnumerator As IEnumeratorConverter(Of IProperty)

#End Region

#Region "Public Properties"

    Public ReadOnly Property Names As IList(Of String)
      Get
        Try
          Return GetPropertyNames()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try

      End Get
    End Property

#End Region

#Region "Private Properties"

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

#Region "Overloads"

    Public Overloads Function Contains(ByVal Name As String) As Boolean
      Try

        Return PropertyExists(Name, False)

        'For Each lobjProperty As ClassificationProperty In Me
        '  If String.Equals(lobjProperty.Name, Name, StringComparison.InvariantCultureIgnoreCase) Then
        '    Return True
        '  ElseIf String.Equals(lobjProperty.SystemName, Name, StringComparison.InvariantCultureIgnoreCase) Then
        '    Return True
        '  ElseIf String.Equals(lobjProperty.PackedName, Name, StringComparison.InvariantCultureIgnoreCase) Then
        '    Return True
        '  End If
        'Next
        'Return False

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("ClassificationProperties::Contains('{0}')", Name))
        Return False
      End Try
    End Function

    Public Shadows Sub Add(ByVal lpProperty As IProperty) Implements System.Collections.Generic.ICollection(Of IProperty).Add
      Try
        If TypeOf lpProperty Is ClassificationProperty Then
          Add(CType(lpProperty, ClassificationProperty))
        Else
          MyBase.Add(New ClassificationProperty(lpProperty))
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub Add(ByVal lpProperty As ECMProperty)
      ' Cast the ECMProperty to a ClassificationProperty
      Try
        MyBase.Add(New ClassificationProperty(lpProperty))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shadows Sub Add(ByVal lpProperty As ClassificationProperty)
      Try
        MyBase.Add(lpProperty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

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

    Public Shadows Sub Add(ByVal lpProperties As ClassificationProperties)
      Try
        For Each lobjProperty As ClassificationProperty In lpProperties
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

    Public Function ToXmlString() As String
      Try
        Return SerializationUtilities.Serializer.Serialize.XmlString(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#Region "Item Overloads"

    'Default Shadows Property Item(ByVal lpName As String) As ClassificationProperty
    '  Get
    '    Try
    '      Dim lobjClassificationProperty As ClassificationProperty
    '      For lintCounter As Integer = 0 To MyBase.Count - 1
    '        lobjClassificationProperty = CType(MyBase.Item(lintCounter), ClassificationProperty)
    '        If lobjClassificationProperty.Name = lpName Then
    '          Return lobjClassificationProperty
    '        End If
    '      Next
    '      'Throw New Exception("There is no Item by the Name '" & lpName & "'.")

    '      Return Nothing

    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, String.Format("ClassificationProperties::Get_Item('{0}')", lpName))
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '    'Throw New InvalidArgumentException
    '  End Get
    '  Set(ByVal value As ClassificationProperty)
    '    Try
    '      Dim lobjClassificationProperty As ClassificationProperty
    '      For lintCounter As Integer = 0 To MyBase.Count - 1
    '        lobjClassificationProperty = CType(MyBase.Item(lintCounter), ClassificationProperty)
    '        If lobjClassificationProperty.Name = lpName Then
    '          MyBase.Item(lintCounter) = value
    '        End If
    '      Next
    '      Throw New Exception("There is no Item by the Name '" & lpName & "'.")
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, String.Format("ClassificationProperties::Set_Item('{0}')", value))
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Set
    'End Property

    '''' <summary>
    '''' Gets or sets the element at the specified index.
    '''' </summary>
    '''' <param name="lpIndex">The zero-based index of the element to get or set.</param>
    '''' <value></value>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Default Shadows Property Item(ByVal lpIndex As Integer) As ClassificationProperty
    '  Get
    '    Return mobjObjects(lpIndex)
    '  End Get
    '  Set(ByVal value As ClassificationProperty)
    '    mobjObjects(lpIndex) = value
    '  End Set
    'End Property

    '''' <summary>
    '''' Gets or sets the element at the specified index.
    '''' </summary>
    '''' <param name="Index">The zero-based index of the element to get or set.</param>
    '''' <value></value>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Default Overridable Property Item(ByVal Index As Integer) As T
    '  Get
    '    Return mobjObjects(Index)
    '  End Get
    '  Set(ByVal value As T)
    '    mobjObjects(Index) = value
    '  End Set
    'End Property

#End Region

#End Region

#Region "Public Methods"

    Public Shadows Sub AddRange(ByVal lpClassificationProperties As ClassificationProperties)
      Try
        Add(lpClassificationProperties)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function PropertyExists(ByVal lpName As String) As Boolean

      ' Does the property exist?
      Try
        Dim lobjFoundClassificationProperty As ClassificationProperty = Nothing
        Return PropertyExists(lpName, lobjFoundClassificationProperty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("ClassificationProperties::PropertyExists('{0}')", lpName))
        Return False
      End Try
    End Function

    ''' <summary>
    ''' Checks to see if the property exists based on comparing the 
    ''' Name of lpProperty to the SystemName of each ClassificationProperty
    ''' </summary>
    ''' <param name="lpProperty"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function PropertyExists(ByVal lpProperty As ECMProperty) As Boolean

      ' Does the property exist?
      Try
        Return PropertyExists(lpProperty, Nothing)
      Catch ex As Exception
        If lpProperty IsNot Nothing AndAlso lpProperty.Name IsNot Nothing Then
          ApplicationLogging.LogException(ex, String.Format("ClassificationProperties::PropertyExists('{0}')", lpProperty.Name))
        Else
          ApplicationLogging.LogException(ex, "ClassificationProperties::PropertyExists()")
        End If
        Return False
      End Try
    End Function

    Public Function PropertyExists(ByVal lpName As String,
                                   ByRef lpFoundProperty As ClassificationProperty) As Boolean

      ' Does the property exist?
      Try
        'Dim lobjProperty As ClassificationProperty = Item(lpName)
        'Return True

        Dim list As Object = From lobjProperty In Items Where
          lobjProperty.Name = lpName Or lobjProperty.SystemName = lpName Select lobjProperty

        For Each lobjProperty As ClassificationProperty In list
          lpFoundProperty = lobjProperty
          Return True
        Next

        Return False

        'For Each lobjProperty As ClassificationProperty In Me
        '  If StrComp(lobjProperty.Name, lpName, CompareMethod.Binary) = 0 Then
        '    lpFoundProperty = lobjProperty
        '    Return True
        '  End If
        'Next

        'Return False

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("ClassificationProperties::PropertyExists('{0}')", lpName))
        Return False
      End Try
    End Function

    Public Overridable Function PropertyExists(ByVal lpName As String, ByVal lpCaseSensitive As Boolean) As Boolean Implements IProperties.PropertyExists
      ' Does the property exist?
      Try

        Dim list As Object = From lobjProperty In Items Where
        (String.Compare(lobjProperty.Name, lpName, Not lpCaseSensitive) = 0) Or
        (String.Compare(lobjProperty.SystemName, lpName, Not lpCaseSensitive) = 0) Select lobjProperty

        For Each lobjProperty As ClassificationProperty In list
          Return True
        Next

        Return False

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("ClassificationProperties::PropertyExists('{0}')", lpName))
        Return False
      End Try
    End Function

    ''' <summary>
    ''' Checks to see if the property exists based on comparing the 
    ''' Name of lpProperty to the SystemName of each ClassificationProperty
    ''' </summary>
    ''' <param name="lpProperty"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function PropertyExists(ByVal lpProperty As ECMProperty,
                                   ByRef lpFoundProperty As ClassificationProperty) As Boolean

      ' Does the property exist?
      Try


        Dim lblnPropertyExists As Boolean

        If Not String.IsNullOrEmpty(lpProperty.SystemName) Then
          ' Check based on the system name only, otherwise we could get a true return that is not actually valid for loading
          lblnPropertyExists = PropertyExists(lpProperty.SystemName, lpFoundProperty)
        End If

        If lblnPropertyExists = False Then
          lblnPropertyExists = PropertyExists(lpProperty.Name, lpFoundProperty)
        End If

        Return lblnPropertyExists

      Catch ex As Exception
        If lpProperty IsNot Nothing AndAlso lpProperty.Name IsNot Nothing Then
          ApplicationLogging.LogException(ex, String.Format("ClassificationProperties::PropertyExists('{0}')", lpProperty.Name))
        Else
          ApplicationLogging.LogException(ex, "ClassificationProperties::PropertyExists()")
        End If
        Return False
      End Try
    End Function

    Public Function GetPropertyNames() As List(Of String)
      Try
        Dim lobjPropertyNames As New List(Of String)

        For Each lobjProperty As IProperty In Me
          lobjPropertyNames.Add(lobjProperty.Name)
        Next

        lobjPropertyNames.Sort()

        Return lobjPropertyNames

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Friend Sub GetSubscribedClasses(lpDocumentClasses As DocumentClasses)
      Try
        For Each lobjProperty As ClassificationProperty In Me
          lobjProperty.SubscribedClasses = lobjProperty.GetSubscribedClasses(lpDocumentClasses)
        Next
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "INamedItems Implementation"

    Public Shadows Sub Add(ByVal item As Object) Implements INamedItems.Add
      Try
        If TypeOf (item) Is ClassificationProperty Then
          Add(DirectCast(item, ClassificationProperty))
        ElseIf TypeOf (item) Is ClassificationProperties Then
          For Each lobjProperty As ClassificationProperty In item
            Add(DirectCast(lobjProperty, ClassificationProperty))
          Next
        Else
          Throw New ArgumentException(
            String.Format("Invalid object type.  Item type of ClassificationProperty expected instead of type {0}",
                          item.GetType.Name), "item")
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public ReadOnly Property ItemById(ByVal id As String) As ClassificationProperty
      Get
        For Each lobjClassificationProperty As ClassificationProperty In Me
          If String.Equals(lobjClassificationProperty.ID, id, StringComparison.InvariantCultureIgnoreCase) Then
            Return lobjClassificationProperty
          End If
        Next
        Return Nothing
      End Get
    End Property

    'Default Overrides Property Item(ByVal index As Integer) As ClassificationProperty
    '	Get
    '		Try
    '			Return MyBase.Item(index)
    '		Catch ex As Exception
    '			ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '			'   Re-throw the exception to the caller
    '			Throw
    '		End Try
    '	End Get
    '	Set(value As ClassificationProperty)
    '		Try
    '			MyBase.Item(index) = value
    '		Catch ex As Exception
    '			ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '			'   Re-throw the exception to the caller
    '			Throw
    '		End Try
    '	End Set
    'End Property

    'Default Overrides Property Item(ByVal name As String) As ClassificationProperty
    '	Get
    '		Try
    '			Return ItemByName(name)
    '		Catch ex As Exception
    '			ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '			'   Re-throw the exception to the caller
    '			Throw
    '		End Try
    '	End Get
    '	Set(value As ClassificationProperty)
    '		Try
    '			ItemByName(name) = value
    '		Catch ex As Exception
    '			ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '			'   Re-throw the exception to the caller
    '			Throw
    '		End Try
    '	End Set
    'End Property

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public Property ItemByName(ByVal name As String) As Object Implements INamedItems.ItemByName
      Get
        Try
          Dim lobjItem As Object = MyBase.Item(name)
          If lobjItem IsNot Nothing Then
            Return lobjItem
          Else
            For Each lobjProperty As ClassificationProperty In MyBase.Items
              If String.Equals(lobjProperty.SystemName, name, StringComparison.InvariantCultureIgnoreCase) Then
                Return lobjProperty
              ElseIf String.Equals(lobjProperty.PackedName, name, StringComparison.InvariantCultureIgnoreCase) Then
                Return lobjProperty
              End If
            Next
            Return Nothing
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
        Return MyBase.Item(name)
      End Get
      Set(ByVal value As Object)
        MyBase.Item(name) = value
      End Set
    End Property

    <XmlIgnore()>
    Property IItem(ByVal Name As String) As IProperty Implements IProperties.Item
      Get
        Try
          Return ItemByName(Name)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As IProperty)
        Try
          ItemByName(Name) = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Shadows Function Remove(ByVal lpItemName As String) As Boolean Implements INamedItems.Remove
      Try
        If Contains(lpItemName) Then
          Return MyBase.Remove(Item(lpItemName))
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

#Region "IProperties Implementation"

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
        Return Contains(item.Name)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Sub CopyTo(ByVal array() As IProperty, ByVal arrayIndex As Integer) Implements System.Collections.Generic.ICollection(Of IProperty).CopyTo
      Try
        CopyTo(array, arrayIndex)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of IProperty) Implements System.Collections.Generic.IEnumerable(Of IProperty).GetEnumerator
      Try
        Return IPropertyEnumerator.GetEnumerator()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shadows Function Remove(ByVal item As IProperty) As Boolean Implements System.Collections.Generic.ICollection(Of IProperty).Remove
      Try
        Return Remove(item.Name)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

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

#End Region

  End Class

End Namespace
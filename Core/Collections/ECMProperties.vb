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

  Public Class ECMProperties
    Inherits CCollection(Of ECMProperty)

#Region "Public Methods"

    ''' <summary>
    ''' Gets or sets the property using the property name
    ''' </summary>
    ''' <param name="Name">The name of the property</param>
    ''' <value></value>
    ''' <returns>An ECMProperty reference if the property name
    ''' is found, otherwise returns null.</returns>
    ''' <remarks></remarks>
    <XmlElement("Property", GetType(ECMProperty))>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Default Shadows Property Item(ByVal Name As String) As ECMProperty
      Get
        Dim lobjProperty As ECMProperty
        ' First look for the class by Name
        For lintCounter As Integer = 0 To MyBase.Count - 1
          lobjProperty = CType(MyBase.Item(lintCounter), ECMProperty)
          If String.Equals(lobjProperty.Name, Name, StringComparison.InvariantCultureIgnoreCase) Then
            Return lobjProperty
          End If
        Next

        ' Next look for the class by SystemName
        For lintCounter As Integer = 0 To MyBase.Count - 1
          lobjProperty = CType(MyBase.Item(lintCounter), ECMProperty)
          If String.Equals(lobjProperty.SystemName, Name, StringComparison.InvariantCultureIgnoreCase) Then
            Return lobjProperty
          End If
        Next

        ' Next look for the class by PackedName
        For lintCounter As Integer = 0 To MyBase.Count - 1
          lobjProperty = CType(MyBase.Item(lintCounter), ECMProperty)
          If String.Equals(lobjProperty.PackedName, Name, StringComparison.InvariantCultureIgnoreCase) Then
            Return lobjProperty
          End If
        Next

        'Throw New Exception("There is no Item by the Name '" & Name & "'.")
        'Throw New InvalidArgumentException

        ' If the property can't be found, just return Nothing
        Return Nothing

      End Get
      Set(ByVal value As ECMProperty)
        Dim lobjProperty As ECMProperty
        For lintCounter As Integer = 0 To MyBase.Count - 1
          lobjProperty = CType(MyBase.Item(lintCounter), ECMProperty)
          If String.Equals(lobjProperty.Name, Name, StringComparison.InvariantCultureIgnoreCase) Then
            MyBase.Item(lintCounter) = value
          End If
        Next
        Throw New Exception("There is no Item by the Name '" & Name & "'.")
      End Set
    End Property

    Public Overridable Function PropertyExists(ByVal lpName As String,
                                           ByVal lpCaseSensitive As Boolean,
                                           ByRef lpFoundProperty As ECMProperty) As Boolean

      ' Does the property exist?
      Try
        'ApplicationLogging.WriteLogEntry("Enter Properties::PropertyExists", TraceEventType.Verbose)
        'For Each lobjProperty As ECMProperty In Me

        lpFoundProperty = Nothing

        If Me.Count = 0 Then
          Return False
        End If

        If lpCaseSensitive Then


          lpFoundProperty = MyBase.FirstOrDefault(Function(IProperty) IProperty.Name = lpName)

          If lpFoundProperty IsNot Nothing Then
            Return True
          End If

          lpFoundProperty = MyBase.FirstOrDefault(Function(IProperty) IProperty.SystemName = lpName)

          If lpFoundProperty IsNot Nothing Then
            Return True
          End If

        Else

          lpFoundProperty = MyBase.FirstOrDefault(Function(IProperty) String.Equals(IProperty.Name, lpName, StringComparison.InvariantCultureIgnoreCase))

          If lpFoundProperty IsNot Nothing Then
            Return True
          End If

          lpFoundProperty = MyBase.FirstOrDefault(Function(IProperty) String.Equals(IProperty.SystemName, lpName, StringComparison.InvariantCultureIgnoreCase))

          If lpFoundProperty IsNot Nothing Then
            Return True
          End If

        End If

        Return False

        'For Each lobjProperty As IProperty In Me
        '  If lpCaseSensitive Then
        '    ' Check the name
        '    If String.Equals(lobjProperty.Name, lpName) Then
        '      lpFoundProperty = lobjProperty
        '      Return True
        '    End If
        '    ' Check the system name
        '    If String.Equals(lobjProperty.SystemName, lpName) Then
        '      lpFoundProperty = lobjProperty
        '      Return True
        '    End If
        '  Else
        '    ' Check the name
        '    If String.Equals(lobjProperty.Name, lpName, StringComparison.InvariantCultureIgnoreCase) Then
        '      lpFoundProperty = lobjProperty
        '      Return True
        '    End If
        '    ' Check the system name
        '    If String.Equals(lobjProperty.SystemName, lpName, StringComparison.InvariantCultureIgnoreCase) Then
        '      lpFoundProperty = lobjProperty
        '      Return True
        '    End If
        '  End If
        'Next

        'lpFoundProperty = Nothing
        'Return False

      Catch ex As Exception
        ApplicationLogging.LogException(ex, "Properties::PropertyExists")
        Return False
      Finally
        'ApplicationLogging.WriteLogEntry("Exit Properties::PropertyExists", TraceEventType.Verbose)
      End Try
    End Function

    Public Overridable Function PropertyExists(ByVal lpName As String, ByVal lpCaseSensitive As Boolean) As Boolean

      ' Does the property exist?
      Try
        Return PropertyExists(lpName, lpCaseSensitive, Nothing)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, "Properties::PropertyExists")
        Return False
      Finally
        'ApplicationLogging.WriteLogEntry("Exit Properties::PropertyExists", TraceEventType.Verbose)
      End Try
    End Function

    Public Overridable Function PropertyExists(ByVal lpName As String) As Boolean

      ' Does the property exist?
      Try
        'ApplicationLogging.WriteLogEntry("Enter Properties::PropertyExists", TraceEventType.Verbose)
        'Return PropertyExists(lpName, True)
        Return PropertyExists(lpName, True, Nothing)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, "Properties::PropertyExists")
        Return False
      Finally
        'ApplicationLogging.WriteLogEntry("Exit Properties::PropertyExists", TraceEventType.Verbose)
      End Try
    End Function

    ''' <summary>
    ''' Removes all read only properties from the collection
    ''' </summary>
    ''' <param name="lpDocumentClass"></param>
    ''' <remarks></remarks>
    Public Sub RemoveReadOnlyProperties(ByVal lpDocumentClass As DocumentClass)
      Try

        Dim lobjClassificationProperty As ClassificationProperty = Nothing
        Dim lstrPropertyName As String = String.Empty

        For lintPropertyCounter As Integer = Count - 1 To 0 Step -1
          lstrPropertyName = Item(lintPropertyCounter).Name
          'lobjClassificationProperty = lpDocumentClass.Properties(Item(lintPropertyCounter).Name)
          lobjClassificationProperty = lpDocumentClass.Properties.ItemByName(lstrPropertyName)
          If lobjClassificationProperty IsNot Nothing Then
            If lobjClassificationProperty.Settability = ClassificationProperty.SettabilityEnum.READ_ONLY Then
              Remove(Item(lintPropertyCounter))
              ApplicationLogging.WriteLogEntry(String.Format("Removed read only property '{0}' from collection", lobjClassificationProperty.Name),
                   TraceEventType.Information, 4256)
            End If
          End If
        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Gets all the properties in the collection of the specified type.
    ''' </summary>
    ''' <param name="lpPropertyType">The property type to select</param>
    ''' <returns></returns>
    ''' <remarks>If no properties exist in the collection for the specified 
    ''' type, an empty collection is returned.</remarks>
    Public Function PropertiesByType(lpPropertyType As PropertyType) As ECMProperties
      Try
        Dim lobjReturnProperties As New ECMProperties
        Dim list As Object = From lobjProperty As ECMProperty In Me Where
               lobjProperty.Type = lpPropertyType Select lobjProperty
        lobjReturnProperties.AddRange(list)
        Return lobjReturnProperties
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shadows Sub Sort()
      Try
        MyBase.Sort(New PropertyComparer)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace
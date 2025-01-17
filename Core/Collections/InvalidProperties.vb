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

  Public Class InvalidProperties
    Inherits CCollection(Of InvalidProperty)

#Region "Public Methods"

    ''' <summary>
    ''' Gets or sets the property using the property name
    ''' </summary>
    ''' <param name="Name">The name of the property</param>
    ''' <value></value>
    ''' <returns>An ECMProperty reference if the property name
    ''' is found, otherwise returns null.</returns>
    ''' <remarks></remarks>
    <XmlElement("Property", GetType(InvalidProperty))>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Default Shadows Property Item(ByVal Name As String) As InvalidProperty
      Get
        Try
          Dim lobjProperty As InvalidProperty
          ' First look for the class by Name
          For lintCounter As Integer = 0 To MyBase.Count - 1
            lobjProperty = CType(MyBase.Item(lintCounter), InvalidProperty)
            If String.Equals(lobjProperty.Name, Name, StringComparison.InvariantCultureIgnoreCase) Then
              Return lobjProperty
            End If
          Next

          ' Next look for the class by SystemName
          For lintCounter As Integer = 0 To MyBase.Count - 1
            lobjProperty = CType(MyBase.Item(lintCounter), InvalidProperty)
            If String.Equals(lobjProperty.SystemName, Name, StringComparison.InvariantCultureIgnoreCase) Then
              Return lobjProperty
            End If
          Next

          ' Next look for the class by PackedName
          For lintCounter As Integer = 0 To MyBase.Count - 1
            lobjProperty = CType(MyBase.Item(lintCounter), InvalidProperty)
            If String.Equals(lobjProperty.PackedName, Name, StringComparison.InvariantCultureIgnoreCase) Then
              Return lobjProperty
            End If
          Next

          'Throw New Exception("There is no Item by the Name '" & Name & "'.")
          'Throw New InvalidArgumentException

          ' If the property can't be found, just return Nothing
          Return Nothing
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As InvalidProperty)
        Try
          Dim lobjProperty As InvalidProperty
          For lintCounter As Integer = 0 To MyBase.Count - 1
            lobjProperty = CType(MyBase.Item(lintCounter), InvalidProperty)
            If String.Equals(lobjProperty.Name, Name, StringComparison.InvariantCultureIgnoreCase) Then
              MyBase.Item(lintCounter) = value
            End If
          Next
          Throw New Exception("There is no Item by the Name '" & Name & "'.")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Default Shadows Property Item(ByVal index As Integer) As InvalidProperty
      Get
        Try
          Return MyBase.Item(index)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As InvalidProperty)
        Try
          MyBase.Item(index) = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Overridable Function PropertyExists(ByVal lpName As String,
                                           ByVal lpCaseSensitive As Boolean,
                                           ByRef lpFoundProperty As InvalidProperty) As Boolean

      ' Does the property exist?
      Try
        ApplicationLogging.WriteLogEntry("Enter Properties::PropertyExists", TraceEventType.Verbose)
        ''Dim lobjCurrentProperty As InvalidProperty = Nothing

        ''For lintPropertyCounter As Integer = 0 To Count - 1
        ''  lobjCurrentProperty = Items(lintPropertyCounter)
        ''  If lpCaseSensitive Then
        ''    ' Check the name
        ''    If String.Equals(lobjCurrentProperty.Name, lpName) Then
        ''      lpFoundProperty = lobjCurrentProperty
        ''      Return True
        ''    End If
        ''    ' Check the system name
        ''    If String.Equals(lobjCurrentProperty.SystemName, lpName) Then
        ''      lpFoundProperty = lobjCurrentProperty
        ''      Return True
        ''    End If
        ''  Else
        ''    ' Check the name
        ''    If String.Equals(lobjCurrentProperty.Name, lpName, StringComparison.InvariantCultureIgnoreCase) Then
        ''      lpFoundProperty = lobjCurrentProperty
        ''      Return True
        ''    End If
        ''    ' Check the system name
        ''    If String.Equals(lobjCurrentProperty.SystemName, lpName, StringComparison.InvariantCultureIgnoreCase) Then
        ''      lpFoundProperty = lobjCurrentProperty
        ''      Return True
        ''    End If
        ''  End If
        ''Next
        For Each lobjProperty As InvalidProperty In Me
          If lpCaseSensitive Then
            ' Check the name
            If String.Equals(lobjProperty.Name, lpName) Then
              lpFoundProperty = lobjProperty
              Return True
            End If
            ' Check the system name
            If String.Equals(lobjProperty.SystemName, lpName) Then
              lpFoundProperty = lobjProperty
              Return True
            End If
          Else
            ' Check the name
            If String.Equals(lobjProperty.Name, lpName, StringComparison.InvariantCultureIgnoreCase) Then
              lpFoundProperty = lobjProperty
              Return True
            End If
            ' Check the system name
            If String.Equals(lobjProperty.SystemName, lpName, StringComparison.InvariantCultureIgnoreCase) Then
              lpFoundProperty = lobjProperty
              Return True
            End If
          End If
        Next

        lpFoundProperty = Nothing
        Return False

      Catch ex As Exception
        ApplicationLogging.LogException(ex, "Properties::PropertyExists")
        Return False
      Finally
        ApplicationLogging.WriteLogEntry("Exit Properties::PropertyExists", TraceEventType.Verbose)
      End Try
    End Function

    Public Overridable Function PropertyExists(ByVal lpName As String) As Boolean

      ' Does the property exist?
      Try
        ApplicationLogging.WriteLogEntry("Enter Properties::PropertyExists", TraceEventType.Verbose)
        'Return PropertyExists(lpName, True)
        Return PropertyExists(lpName, True, Nothing)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, "Properties::PropertyExists")
        Return False
      Finally
        ApplicationLogging.WriteLogEntry("Exit Properties::PropertyExists", TraceEventType.Verbose)
      End Try
    End Function

    'Public Shadows Sub Add(ByVal item As IProperty)
    '  Try
    '    If TypeOf item Is InvalidProperty Then
    '      If PropertyExists(item.Name) = False Then
    '        MyBase.Add(DirectCast(item, InvalidProperty))
    '      End If
    '    Else
    '      ApplicationLogging.WriteLogEntry( _
    '        String.Format("Unable to add property '{0}' to collection, it must inherit from InvalidProperty.", item.Name), _
    '        TraceEventType.Warning, 62914)
    '    End If
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    Public Shadows Sub Add(ByVal item As ECMProperty, ByVal scope As InvalidProperty.InvalidPropertyScope)
      Try
        MyBase.Add(New InvalidProperty(item, scope))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace

'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Text
Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Transformations

  Public Class LookupProperties
    Inherits CCollection(Of LookupProperty)

#Region "Public Methods"

    Public Overloads Sub Add(ByVal Properties As LookupProperties)
      Try
        For Each lobjLookupProperty As LookupProperty In Properties
          If Me.PropertyExists(lobjLookupProperty.Name) Then
            ApplicationLogging.WriteLogEntry(
              String.Format("The property '{0}' is already defined in the collection, the property was not added.",
                            lobjLookupProperty.Name), TraceEventType.Warning, 9823)
          Else
            Add(lobjLookupProperty)
          End If
        Next
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Gets or sets the property using the property name
    ''' </summary>
    ''' <param name="Name">The name of the property</param>
    ''' <value></value>
    ''' <returns>An LookupProperty reference if the property name
    ''' is found, otherwise returns null.</returns>
    ''' <remarks></remarks>
    <XmlElement("Property", GetType(LookupProperty))>
    Default Shadows Property Item(ByVal Name As String) As LookupProperty
      Get
        Try
          'Dim lobjProperty As LookupProperty
          '' First look for the class by Name
          'For lintCounter As Integer = 0 To MyBase.Count - 1
          '  lobjProperty = CType(MyBase.Item(lintCounter), LookupProperty)
          '  If String.Compare(lobjProperty.Name, Name, True) = 0 Then
          '    Return lobjProperty
          '  End If
          'Next

          '' If the property can't be found, just return Nothing
          'Return Nothing

          Return MyBase.Item(Name)

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get

      Set(ByVal value As LookupProperty)
        Try
          Dim lobjProperty As LookupProperty
          For lintCounter As Integer = 0 To MyBase.Count - 1
            lobjProperty = CType(MyBase.Item(lintCounter), LookupProperty)
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

    <XmlElement("Property", GetType(LookupProperty))> Default Shadows Property Item(ByVal Index As Integer) As LookupProperty
      Get
        Return MyBase.Item(Index)
      End Get
      Set(ByVal value As LookupProperty)
        MyBase.Item(Index) = value
      End Set
    End Property

    Public Overloads Function Delete(ByVal lpName As String) As Boolean
      Try
        Return Remove(Item(lpName))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Function Remove(ByVal lpName As String) As Boolean
      Try
        Return Remove(Item(lpName))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Function PropertyExists(ByVal lpName As String,
                                               ByVal lpCaseSensitive As Boolean,
                                               ByRef lpFoundProperty As LookupProperty) As Boolean

      ' Does the property exist?
      Try
        ApplicationLogging.WriteLogEntry("Enter LookupProperties::PropertyExists", TraceEventType.Verbose)
        For Each lobjProperty As LookupProperty In Me
          If String.Compare(lobjProperty.Name, lpName, Not lpCaseSensitive) = 0 Then
            lpFoundProperty = lobjProperty
            Return True
          End If
        Next

        lpFoundProperty = Nothing
        Return False

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return False
      Finally
        ApplicationLogging.WriteLogEntry("Exit LookupProperties::PropertyExists", TraceEventType.Verbose)
      End Try
    End Function

    Public Overridable Function PropertyExists(ByVal lpName As String, ByVal lpCaseSensitive As Boolean) As Boolean

      ' Does the property exist?
      Try
        Return PropertyExists(lpName, lpCaseSensitive, Nothing)
        'ApplicationLogging.WriteLogEntry("Enter Properties::PropertyExists", TraceEventType.Verbose)
        'For Each lobjProperty As LookupProperty In Me
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
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return False
      Finally
        ApplicationLogging.WriteLogEntry("Exit LookupProperties::PropertyExists", TraceEventType.Verbose)
      End Try
    End Function

    Public Overridable Function PropertyExists(ByVal lpName As String) As Boolean

      ' Does the property exist?
      Try
        ApplicationLogging.WriteLogEntry("Enter LookupProperties::PropertyExists", TraceEventType.Verbose)
        'Return PropertyExists(lpName, True)
        Return PropertyExists(lpName, True, Nothing)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return False
      Finally
        ApplicationLogging.WriteLogEntry("Exit Properties::PropertyExists", TraceEventType.Verbose)
      End Try
    End Function

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

        For Each lobjProperty As LookupProperty In Me
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

  End Class

End Namespace
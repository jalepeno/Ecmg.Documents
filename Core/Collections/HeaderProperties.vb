'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports System.Xml.Serialization
Imports Documents.Utilities

Namespace Core
  ''' <exclude/>
  <Microsoft.VisualBasic.HideModuleName()>
  <Serializable()>
  Public Class HeaderProperties
    Inherits CCollection(Of HeaderProperty)

#Region "Public Methods"

    <XmlElement("Property", GetType(HeaderProperty))> Default Shadows Property Item(ByVal Name As String) As HeaderProperty
      Get
        Dim lobjProperty As HeaderProperty
        ' First look for the class by Name
        For lintCounter As Integer = 0 To MyBase.Count - 1
          lobjProperty = CType(MyBase.Item(lintCounter), HeaderProperty)
          If lobjProperty.Name.ToUpper = Name.ToUpper Then
            Return lobjProperty
          End If
        Next

        ' Next look for the class by PackedName
        For lintCounter As Integer = 0 To MyBase.Count - 1
          lobjProperty = CType(MyBase.Item(lintCounter), HeaderProperty)
          If lobjProperty.PackedName.ToUpper = Name.ToUpper Then
            Return lobjProperty
          End If
        Next

        Throw New Exception("There is no Item by the Name '" & Name & "'.")
        'Throw New InvalidArgumentException
      End Get
      Set(ByVal value As HeaderProperty)
        If Helper.IsDeserializationBasedCall Then
          Dim lobjProperty As HeaderProperty
          For lintCounter As Integer = 0 To MyBase.Count - 1
            lobjProperty = CType(MyBase.Item(lintCounter), HeaderProperty)
            If lobjProperty.Name = Name Then
              MyBase.Item(lintCounter) = value
            End If
          Next
          Throw New Exception("There is no Item by the Name '" & Name & "'.")
        Else
          Throw New InvalidOperationException("Although 'Item' is a public property, set operations are not allowed.  Treat as read-only.")
        End If
      End Set
    End Property

    <XmlElement("Property", GetType(HeaderProperty))> Default Shadows Property Item(ByVal Index As Integer) As HeaderProperty
      Get
        Return MyBase.Item(Index)
      End Get
      Set(ByVal value As HeaderProperty)
        MyBase.Item(Index) = value
      End Set
    End Property

    Public Overloads Function Delete(ByVal lpName As String) As Boolean
      Return Remove(Item(lpName))
    End Function

    Public Overridable Function PropertyExists(ByVal lpName As String) As Boolean

      ' Does the property exist?
      Try
        ApplicationLogging.WriteLogEntry("Enter Properties::PropertyExists", TraceEventType.Verbose)
        For Each lobjProperty As HeaderProperty In Me
          If lobjProperty.Name = lpName Then
            Return True
          End If
        Next
        Return False
      Catch ex As Exception
        ApplicationLogging.LogException(ex, "Properties::PropertyExists")
        Return False
      Finally
        ApplicationLogging.WriteLogEntry("Exit Properties::PropertyExists", TraceEventType.Verbose)
      End Try
    End Function

#End Region

  End Class
End Namespace
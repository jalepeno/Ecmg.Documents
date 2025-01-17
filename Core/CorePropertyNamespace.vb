'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel

#End Region

Namespace Core

#Region "Public Interfaces"

  Public Interface IProperty

    Property Type As PropertyType
    Property Cardinality As Cardinality
    Property Name As String
    Property Description As String
    Property DisplayName As String
    Property SystemName As String
    Property DefaultValue As Object
    ReadOnly Property HasValue As Boolean
    'Property Value As Object
    ReadOnly Property Persistent As Boolean
    Property Value As Object
    Property Values As Object

    ReadOnly Property HasStandardValues As Boolean

    ReadOnly Property StandardValues As IEnumerable

    ''' <summary>
    ''' Clears current value(s) from the property.
    ''' </summary>
    ''' <remarks>
    ''' Clears both the Value and Values properties
    ''' </remarks>
    Sub Clear()
    Function ToDebugString() As String

  End Interface

  Public Interface IRepositoryKeyList

    Property SelectedKeys As List(Of String)

  End Interface

  Public Interface IProperties
    Inherits ICollection(Of IProperty)
#If SILVERLIGHT <> 1 Then
    Inherits ICloneable
#End If

    Property Item(ByVal Name As String) As IProperty
    Function PropertyExists(ByVal lpName As String, ByVal lpCaseSensitive As Boolean) As Boolean
    Sub Replace(ByVal lpName As String, ByVal lpNewProperty As IProperty)
    Overloads Sub Add(ByVal lpProperties As IProperties)

  End Interface

  Public Interface ISingletonProperty
    Inherits IProperty

    Overloads Property Value As Object

  End Interface

  Public Interface IMultiValuedProperty
    Inherits IProperty

    Overloads Property Values As Values

  End Interface

#End Region

#Region "Public Enumerations"

  ''' <summary>
  ''' Defines the data type for a property value.
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum PropertyType
    <Description("Undefined")>
    ecmUndefined = 0
    <Description("Binary")>
    ecmBinary = 1
    <Description("Boolean")>
    ecmBoolean = 2
    <Description("Date")>
    ecmDate = 3
    <Description("Double")>
    ecmDouble = 4
    <Description("Guid")>
    ecmGuid = 5
    <Description("Long")>
    ecmLong = 6
    <Description("Object")>
    ecmObject = 7
    <Description("String")>
    ecmString = 8
    <Description("Uri")>
    ecmUri = 9
    <Description("Xml")>
    ecmXml = 10
    <Description("Html")>
    ecmHtml = 11
    <Description("Enumeration")>
    ecmEnum = 12
    <Description("Value Map")>
    ecmValueMap = 13
  End Enum

  ''' <summary>The object scope to which the property is defined</summary>
  Public Enum PropertyScope
    ''' <summary>The property is defined at the Document level</summary>
    DocumentProperty = 0
    ''' <summary>The property is defined at the Version level</summary>
    VersionProperty = 1
    ''' <summary>The property may be defined at either the Document or Version level</summary>
    BothDocumentAndVersionProperties = 3
    ''' <summary>The property is defined at the Content level</summary>
    ContentProperty = 4
    ''' <summary>The property may be defined at any level</summary>
    AllProperties = 5
  End Enum

  ''' <summary>Defines the cardinality of an ECMProperty object.</summary>
  Public Enum Cardinality
    <Description("Multi-valued")>
    ecmMultiValued = 0
    <Description("Single-valued")>
    ecmSingleValued = 1
  End Enum

  ''' <summary>
  ''' Determines any special user interface components to be used with property.
  ''' </summary>
  Public Enum DisplayType
    ''' <summary>Use standard display type (default).</summary>
    Standard = 0
    ''' <summary>Treat this as a password property.</summary>
    Password = 1
    ''' <summary>Used for binding to an OLE DB dialog.</summary>
    OleDbConnectionString = 2
  End Enum

#End Region

End Namespace

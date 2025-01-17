'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization

#End Region

Namespace Annotations.Security

  <Serializable()>
  Public Class Permission

#Region "Public Properties"

    <XmlElement("Create")>
    Public Property CanCreate() As Boolean

    <XmlElement("ViewContent")>
    Public Property CanViewContent() As Boolean

    <XmlElement("ViewMetadata")>
    Public Property CanViewMetadata() As Boolean

    <XmlElement("ModifyContent")>
    Public Property CanModifyContent() As Boolean

    <XmlElement("ModifyMetadata")>
    Public Property CanModifyMetadata() As Boolean

    <XmlElement("DeleteContent")>
    Public Property CanDeleteContent() As Boolean

    <XmlElement("DeleteMetadata")>
    Public Property CanDeleteMetadata() As Boolean

    <XmlElement("ControlContentAccess")>
    Public Property CanControlContentAccess() As Boolean

    <XmlElement("ControlMetadataAccess")>
    Public Property CanControlMetadataAccess() As Boolean

    <XmlElement("TakeOwnership")>
    Public Property CanTakeOwnership() As Boolean
#End Region

  End Class
End Namespace
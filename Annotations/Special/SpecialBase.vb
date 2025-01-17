'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports System.Xml.Serialization

Namespace Annotations.Special

  ''' <summary>
  ''' Marker class for potentially non-portable annotation types
  ''' </summary>
  <Serializable()>
  <XmlInclude(GetType(ContentAnnotation))>
  <XmlInclude(GetType(StampAnnotation))>
  <XmlInclude(GetType(StickyNoteAnnotation))>
  Public MustInherit Class SpecialBase
    Inherits Annotation

  End Class
End Namespace
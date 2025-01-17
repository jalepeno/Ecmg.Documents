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

Namespace Annotations.Text

  ''' <summary>
  ''' A textual markup element.  Intend to replace with a standardized format.
  ''' </summary>
  <Serializable()>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class TextMarkup

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the text of the annotation.
    ''' </summary>
    ''' <value>The text of the annotation.</value>
    <XmlAttribute()>
    Public Property Text() As String

    ''' <summary>
    ''' Gets or sets the text rotation.
    ''' </summary>
    ''' <value>The text rotation.</value>
    <XmlAttribute()>
    Public Property TextRotation() As Single

    ''' <summary>
    ''' Gets or sets the font.
    ''' </summary>
    ''' <value>The font.</value>
    Public Property Font() As FontInfo

#End Region

#Region "Private Methods"

    Protected Friend Overridable Function DebuggerIdentifier() As String
      Try
        If Text IsNot Nothing AndAlso Text.Length > 0 Then
          Return Text
        Else
          Return "Text value not set"
        End If
      Catch ex As System.Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace

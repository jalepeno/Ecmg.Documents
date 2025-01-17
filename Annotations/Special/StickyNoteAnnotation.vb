'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Annotations.Text
Imports Documents.Utilities

#End Region

Namespace Annotations.Special

  ''' <summary>
  ''' Provides a representation of a sticky note without having to draw it using other annotations.
  ''' </summary>
  <Serializable()>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class StickyNoteAnnotation
    Inherits SpecialBase

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the note order.
    ''' </summary>
    ''' <value>The note order.</value>
    Public Property NoteOrder() As Integer

    ''' <summary>
    ''' Gets or sets the text note.
    ''' </summary>
    ''' <value>The text note.</value>
    Public Property TextNote As TextMarkup

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Initializes a new instance of the <see cref="StickyNoteAnnotation" /> class.
    ''' </summary>
    Public Sub New()
      Me.TextNote = New TextMarkup()
    End Sub

#End Region

#Region "Private Methods"

    Protected Overrides Function DebuggerIdentifier() As String
      Try
        If TextNote IsNot Nothing Then
          Return String.Format("{0}: {1}", ClassName, TextNote.DebuggerIdentifier)
        Else
          Return String.Format("{0}: Text value not set", ClassName)
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
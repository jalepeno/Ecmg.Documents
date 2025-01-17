'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Collections.ObjectModel
Imports Documents.Utilities

#End Region

Namespace Annotations.Text

  ''' <summary>
  ''' Represents a text-based annotation and any associated font information.
  ''' </summary>
  <Serializable()>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class TextAnnotation
    Inherits Annotation

#Region "Constructors"

    ''' <summary>
    ''' Initializes a new instance of the <see cref="TextAnnotation"/> class.
    ''' </summary>
    Public Sub New()
      Try
        Me.Angle = 0.0
        Me.TextMarkups = New Collection(Of TextMarkup)()
      Catch ex As System.Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the collection of text markups.
    ''' </summary>
    ''' <value>The collection of text markups.</value>
    Public Property TextMarkups() As Collection(Of TextMarkup)

    ''' <summary>
    ''' Gets or sets the angle.
    ''' </summary>
    ''' <value>The angle.</value>
    Public Property Angle() As Double

#End Region

#Region "Private Methods"

    Protected Overrides Function DebuggerIdentifier() As String
      Try
        If TextMarkups IsNot Nothing Then
          Select Case TextMarkups.Count
            Case 0
              Return String.Format("{0}: Text value not set", ClassName)
            Case 1
              Return String.Format("{0}: {1}", ClassName, TextMarkups(0).DebuggerIdentifier)
            Case Is > 1
              Return String.Format("{0}: {1} + {2} others", ClassName, TextMarkups(0).DebuggerIdentifier, TextMarkups.Count - 1)
            Case Else
              Return String.Format("{0}: Text value not set", ClassName)
          End Select
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
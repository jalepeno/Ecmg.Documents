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

  <Serializable()>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class StampAnnotation
    Inherits SpecialBase

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the text element.
    ''' </summary>
    ''' <value>The text element.</value>
    Public Property TextElement() As TextMarkup

#End Region

#Region "Private Methods"

    Protected Overrides Function DebuggerIdentifier() As String
      Dim lobjIdentifierBuilder As New System.Text.StringBuilder

      Try

        If ID IsNot Nothing AndAlso ID.Length > 0 Then
          lobjIdentifierBuilder.AppendFormat("{0}. ", ID)
        End If

        If TextElement IsNot Nothing Then
          lobjIdentifierBuilder.AppendFormat("{0}: {1} ", ClassName, TextElement.DebuggerIdentifier)
        Else
          lobjIdentifierBuilder.AppendFormat("{0}: Text value not set ", ClassName)
        End If

        If AuditEvents IsNot Nothing Then
          lobjIdentifierBuilder.AppendFormat("({0})", AuditEvents.DebuggerIdentifier)
        End If

        Return lobjIdentifierBuilder.ToString

      Catch ex As System.Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
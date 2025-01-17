'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Comparison

  '<DebuggerDisplay("{DebuggerIdentifier(),nq}")> _
  Public Class ComparisonProperties
    Inherits CCollection(Of ComparisonProperty)

#Region "Public Overloads"

    ''' <summary>
    ''' Gets the item matching the name, scope and versionId
    ''' </summary>
    ''' <param name="name"></param>
    ''' <param name="scope"></param>
    ''' <param name="versionId"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Default Overridable Shadows ReadOnly Property Item(ByVal name As String,
                                                       scope As PropertyScope,
                                                       versionId As String) As ComparisonProperty
      Get
        Try

          Dim list As Object = From lobjProperty In Items Where
            lobjProperty.Name = name And lobjProperty.Scope = scope And
            lobjProperty.VersionId = versionId Select lobjProperty

          For Each lobjProperty As ComparisonProperty In list
            Return lobjProperty
          Next

          Return Nothing

          'If list.count > 0 Then
          '  Return list(0)
          'Else
          '  Return Nothing
          'End If

          'Dim lobjReturnProperty As ComparisonProperty = MyBase.Item(name)

          'If lobjReturnProperty IsNot Nothing Then
          '  If lobjReturnProperty.Scope <> scope Then
          '    Return Nothing
          '  End If
          '  If String.Equals(lobjReturnProperty.VersionId, versionId) = False Then
          '    Return Nothing
          '  End If
          'End If

          'Return lobjReturnProperty

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

    'Protected Overridable Function DebuggerIdentifier() As String
    '  Try
    '    Return String.Format("{0} Items", Count)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    'Throw
    '    Return "Comparison Property Collection"
    '  End Try
    'End Function

  End Class

End Namespace
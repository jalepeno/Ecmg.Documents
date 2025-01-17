'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

Namespace Data
  ''' <summary>Collection of QueryTarget objects.</summary>
  Public Class QueryTargets
    Inherits Core.CCollection(Of QueryTarget)

#Region "Overloaded Methods"

    Public Overloads Function Contains(ByVal Name As String) As Boolean
      Try
        'ApplicationLogging.WriteLogEntry("Enter QueryTargets::Contains", TraceEventType.Verbose)

        For Each target As QueryTarget In Me
          If target.Name = Name Then
            Return True
          End If
        Next
        Return False
      Catch ex As Exception
        ApplicationLogging.LogException(ex, "QueryTargets::Contains")
        Return False
      Finally
        'ApplicationLogging.WriteLogEntry("Exit QueryTargets::Contains", TraceEventType.Verbose)
      End Try
    End Function

#End Region

  End Class
End Namespace
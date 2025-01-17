'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

Namespace Data
  Public Class ResultColumns
    Inherits Core.CCollection(Of String)

    Public Overloads Sub Add(lpResultColumn As String)
      Try
        If Not Contains(lpResultColumn) Then
          MyBase.Add(lpResultColumn)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Function Contains(ByVal name As String) As Boolean
      Try

        Dim lobjItem As Object = Item(name, True)

        If lobjItem IsNot Nothing Then
          Return True
        Else
          Return False
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

  End Class
End Namespace

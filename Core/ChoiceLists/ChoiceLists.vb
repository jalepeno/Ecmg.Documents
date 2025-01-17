' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ChoiceLists.vb
'  Description :  [type_description_here]
'  Created     :  6/21/2012 10:10:05 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Core.ChoiceLists

  Public Class ChoiceLists
    Inherits CCollection(Of ChoiceList)

    Public Overrides Function Contains(name As String) As Boolean
      Try
        For Each lobjChoiceList As ChoiceList In Me
          If String.Compare(lobjChoiceList.Name, name) = 0 Then
            Return True
          End If
        Next
        Return False
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Sub Sort()
      Try
        MyBase.Sort(New ChoiceListComparer)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

  End Class

End Namespace
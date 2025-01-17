'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

Namespace Configuration


  Public Class KeyValuePairConnectionStrings
    Inherits Core.CCollection(Of KeyValuePairConnectionString)

    Public Overloads Sub Add(ByVal lpConnectionString As KeyValuePairConnectionString)
      Try
        'If it exists in the collection, replace it
        If (Me.Item(lpConnectionString.Key) IsNot Nothing) Then
          Me.Item(lpConnectionString.Key).Value = lpConnectionString.Value
        Else
          MyBase.Add(lpConnectionString)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Default Shadows Property Item(ByVal key As String) As KeyValuePairConnectionString
      Get
        Try
          For Each lobjKeyValuePair As KeyValuePairConnectionString In Me
            If (lobjKeyValuePair.Key.Equals(key, StringComparison.CurrentCultureIgnoreCase)) Then
              Return lobjKeyValuePair
            End If
          Next

          Return Nothing

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try

      End Get
      Set(ByVal value As KeyValuePairConnectionString)
        Try
          Dim lobjKeyValuePair As KeyValuePair
          For lintCounter As Integer = 0 To MyBase.Count - 1
            lobjKeyValuePair = CType(MyBase.Item(lintCounter), KeyValuePair)
            If lobjKeyValuePair.Key = key Then
              MyBase.Item(lintCounter) = value
              Exit Property
            End If
          Next
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("KeyValuePairs::Set_Item('{0}')", key))
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Default Shadows Property Item(ByVal Index As Integer) As KeyValuePairConnectionString
      Get
        Try
          Return MyBase.Item(Index)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("Profiles::Get_Item('{0}')", Index))
          ' Re-throw the exception to the caller
          Throw
        Finally
        End Try
      End Get
      Set(ByVal value As KeyValuePairConnectionString)
        MyBase.Item(Index) = value
      End Set
    End Property

  End Class

End Namespace

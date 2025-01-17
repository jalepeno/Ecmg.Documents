'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports System.Text
Imports Documents.Utilities

Namespace Configuration

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class KeyValuePairConnectionString
    Inherits KeyValuePair

    Public Sub New()

    End Sub

    ''' <summary>
    ''' Parses out the connection string name and adds it as the key
    ''' </summary>
    ''' <param name="lpConnectionString"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpConnectionString As String)
      MyBase.Key = Helper.GetInfoFromString(lpConnectionString, "Name")
      MyBase.Value = lpConnectionString
    End Sub

#Region "Protected Methods"

    Protected Friend Overrides Function DebuggerIdentifier() As String
      Dim lobjIdentifierBuilder As New StringBuilder
      Try
        If String.IsNullOrEmpty(Key) Then
          lobjIdentifierBuilder.Append("Key not set")
        Else
          lobjIdentifierBuilder.AppendFormat("{0}: {1}", Key, Value)
        End If

        Return lobjIdentifierBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace


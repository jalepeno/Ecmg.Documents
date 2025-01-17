' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  PermissionList.vb
'  Description :  [type_description_here]
'  Created     :  3/16/2012 9:51:18 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Text
Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Security

  Public Class PermissionList
    Inherits CCollection(Of PermissionRight)
    Implements IPermissionList

#Region "IPermissionList Implementation"

    Public Overloads Sub Add(permissionName As String) Implements IPermissionList.Add
      Try
        Dim lenuPermissionRight As Nullable(Of PermissionRight) = [Enum].Parse(GetType(PermissionRight), permissionName, True)

        If ((lenuPermissionRight.HasValue) AndAlso (Me.Contains(lenuPermissionRight.Value) = False)) Then
          MyBase.Add(lenuPermissionRight.Value)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function FromDelimitedList(lpList As String) As IPermissionList Implements IPermissionList.FromDelimitedList
      Try
        Dim lstrRights As String() = lpList.Split(", ")
        Dim lobjReturnList As New PermissionList

        For Each lstrRight As String In lstrRights
          If lstrRight.Length > 0 Then
            lobjReturnList.Add(lstrRight)
          End If
        Next

        Return lobjReturnList

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToDelimitedList() As String Implements IPermissionList.ToDelimitedList
      Try

        If Count = 0 Then
          Return String.Empty
        End If

        Dim lobjListBuilder As New StringBuilder

        For Each lenuRight As PermissionRight In Me
          lobjListBuilder.AppendFormat("{0}, ", lenuRight.ToString())
        Next

        lobjListBuilder.Remove(lobjListBuilder.Length - 2, 2)

        Return lobjListBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
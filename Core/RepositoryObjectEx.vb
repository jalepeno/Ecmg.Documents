'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------


#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Core

  Partial Public Class RepositoryObject
    Implements IComparer

#Region "IComparer Implementation"

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
      Try
        ' Two null objects are equal
        If (x Is Nothing) And (y Is Nothing) Then Return 0

        ' Any non-null object is greater than a null object
        If (x Is Nothing) Then Return 1
        If (y Is Nothing) Then Return -1

        If TypeOf x Is RepositoryObject AndAlso TypeOf y Is RepositoryObject Then
          If DirectCast(x, RepositoryObject).Name < DirectCast(y, RepositoryObject).Name Then
            Return -1
          Else
            Return 1
          End If
        Else
          Throw New ArgumentException("At least one of the parameters is not a RepositoryObject")
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace

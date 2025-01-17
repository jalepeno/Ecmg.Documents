'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Newtonsoft.Json

#End Region

Namespace Utilities

  <Serializable()>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  <JsonConverter(GetType(FileSizeConverter))>
  Partial Public Class FileSize
    Implements IComparable

#Region "Public methods"

#Region "Comparison Methods"

    Public Function GreaterThan(ByVal lpFileSize As FileSize) As Boolean
      Try
        If Bytes.CompareTo(lpFileSize.Bytes) > 0 Then
          Return True
        Else
          Return False
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function LessThan(ByVal lpFileSize As FileSize) As Boolean

      Try
        Return Not GreaterThan(lpFileSize)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Public Overrides Function ToString() As String
      Try
        Return DebuggerIdentifier()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#End Region

#Region "Private Methods"

    Friend Function DebuggerIdentifier() As String

      Dim lstrReturnValue As String = String.Empty

      Try

        If Gigabytes > 1 Then
          lstrReturnValue = String.Format("{0:n} GB", Gigabytes)
        ElseIf Megabytes > 1 Then
          lstrReturnValue = String.Format("{0:n} MB", Megabytes)
        ElseIf Kilobytes > 1 Then
          lstrReturnValue = String.Format("{0:n} KB", Kilobytes)
        Else
          lstrReturnValue = String.Format("{0:n} Bytes", Bytes)
        End If

        Return lstrReturnValue

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IComparable Implementation"

    Public Function CompareTo(ByVal obj As Object) As Integer Implements IComparable.CompareTo

      Try
        If TypeOf obj Is FileSize Then
          Return Bytes.CompareTo(obj.Bytes)
        Else
          Throw New ArgumentException("Object is not a FileSize")
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

#End Region

  End Class

End Namespace

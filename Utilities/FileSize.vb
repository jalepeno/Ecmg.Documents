'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Namespace Utilities

  ''' <summary>Contains information related to the size of a file.</summary>
  Public Class FileSize

#Region "Class Constants"
    Public Const SET_OPERATION_NOT_SUPPORTED As String = "Set operation not supported for FileSize Properties"
#End Region

#Region "Public Enumerations"

    Public Enum FileScale
      Bytes = 0
      KiloBytes = 1
      MegaBytes = 2
      Gigabytes = 3
    End Enum

#End Region

#Region "Class Variables"
    Private mlngFileSizeBytes As ULong
#End Region

#Region "Public Properties"

    Public Property Bytes() As Long
      Get
        Return mlngFileSizeBytes
      End Get
      Set(ByVal Value As Long)
        'Throw New InvalidOperationException(SET_OPERATION_NOT_SUPPORTED)
        mlngFileSizeBytes = Value
      End Set
    End Property

    Public Property Kilobytes() As Double
      Get
        Try
          Return mlngFileSizeBytes / 1024
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal Value As Double)
        'Throw New InvalidOperationException(SET_OPERATION_NOT_SUPPORTED)
      End Set
    End Property

    Public Property Megabytes() As Double
      Get
        Try
          Return mlngFileSizeBytes / 1024 / 1024
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal Value As Double)
        'Throw New InvalidOperationException(SET_OPERATION_NOT_SUPPORTED)
      End Set
    End Property

    Public Property Gigabytes() As Double
      Get
        Try
          Return mlngFileSizeBytes / 1024 / 1024 / 1024
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal Value As Double)
        'Throw New InvalidOperationException(SET_OPERATION_NOT_SUPPORTED)
      End Set
    End Property

#End Region

#Region "Public Operators"

    Public Shared Operator +(ByVal lpFilesize1 As FileSize, ByVal lpFilesize2 As FileSize) As FileSize
      Try
        Return New FileSize(lpFilesize1.Bytes + lpFilesize2.Bytes)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Operator

    Public Shared Operator -(ByVal lpFilesize1 As FileSize, ByVal lpFilesize2 As FileSize) As FileSize
      Try
        Return New FileSize(lpFilesize1.Bytes - lpFilesize2.Bytes)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Operator

    Public Shared Operator *(ByVal lpFilesize1 As FileSize, ByVal lpFilesize2 As FileSize) As FileSize
      Try
        Return New FileSize(lpFilesize1.Bytes * lpFilesize2.Bytes)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Operator

    Public Shared Operator /(ByVal lpFilesize1 As FileSize, ByVal lpFilesize2 As FileSize) As FileSize
      Try
        Return New FileSize(CLng(lpFilesize1.Bytes / lpFilesize2.Bytes))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Operator

    Public Shared Operator /(ByVal lpFilesize1 As FileSize, ByVal lpBytes As Long) As FileSize
      Try
        Return New FileSize(CLng(lpFilesize1.Bytes / lpBytes))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Operator

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpBytes As Long)
      mlngFileSizeBytes = lpBytes
    End Sub

    Public Sub New(ByVal lpFilePath As String)
      Try
        mlngFileSizeBytes = GetFileByteCountFromFilePath(lpFilePath)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Shared Function FromString(lpFileSizeString As String) As FileSize
      Try

        'If Gigabytes > 1 Then
        '  lstrReturnValue = String.Format("{0:n} GB", Gigabytes)
        'ElseIf Megabytes > 1 Then
        '  lstrReturnValue = String.Format("{0:n} MB", Megabytes)
        'ElseIf Kilobytes > 1 Then
        '  lstrReturnValue = String.Format("{0:n} KB", Kilobytes)
        'Else
        '  lstrReturnValue = String.Format("{0:n} Bytes", Bytes)
        'End If

        'Return lstrReturnValue
        Dim lobjFileSize As New FileSize

        Dim lstrNumber As String = lpFileSizeString.Remove(lpFileSizeString.IndexOf(" "))

        If lpFileSizeString.EndsWith(" Bytes") Then
          lobjFileSize.Bytes = CLng(lstrNumber)

        ElseIf lpFileSizeString.EndsWith(" KB") Then
          lobjFileSize.Bytes = Math.Round(CDbl(lstrNumber) * 1024) ' CDbl(lstrNumber)

        ElseIf lpFileSizeString.EndsWith(" MB") Then
          lobjFileSize.Bytes = Math.Round(CDbl(lstrNumber) * 1024 * 1024) ' CDbl(lstrNumber)

        ElseIf lpFileSizeString.EndsWith(" GB") Then
          lobjFileSize.Bytes = Math.Round(CDbl(lstrNumber) * 1024 * 1024 * 1024) ' CDbl(lstrNumber)

        End If

        Return lobjFileSize

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    Private Function GetFileByteCountFromFilePath(ByVal lpFilePath As String) As Long
      Try
        Helper.VerifyFilePath(lpFilePath, True)

        Dim lobjFileInfo As New IO.FileInfo(lpFilePath)
        Return lobjFileInfo.Length

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
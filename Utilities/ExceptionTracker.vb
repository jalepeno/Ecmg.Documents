'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"


#End Region

Namespace Utilities

  Public Class ExceptionTracker

#Region "Class Variables"

    Private Shared mobjInstance As ExceptionTracker
    Private Shared mintReferenceCount As Integer
    Private Shared mobjLastException As Exception
    Private Shared mstrLastExceptionLocation As String = String.Empty

#End Region

#Region "Constructors"

    Private Sub New()
      mintReferenceCount = 0
    End Sub

#End Region

#Region "Singleton Support"

    Public Shared Function Instance() As ExceptionTracker
      Try
        If mobjInstance Is Nothing Then
          mobjInstance = New ExceptionTracker
        End If
        mintReferenceCount += 1
        Return mobjInstance
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Public Properties"

    Public Shared ReadOnly Property LastException As Exception
      Get
        Return mobjLastException
      End Get
    End Property

    Public Shared ReadOnly Property LastExceptionLocation As String
      Get
        Try
          Return mstrLastExceptionLocation
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Public Methods"

    Public Shared Sub Clear()
      Try
        mobjLastException = Nothing
        mstrLastExceptionLocation = String.Empty
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Friend Methods"

    ''' <summary>
    ''' Updates the last exception if it is 
    ''' different than the one currently stored
    ''' </summary>
    ''' <param name="lpException">The new exception</param>
    ''' <param name="lpLocation">The new location</param>
    ''' <remarks></remarks>
    Friend Shared Sub Update(lpException As Exception, lpLocation As String)
      Try
        If mobjLastException Is Nothing Then
          mobjLastException = lpException
          mstrLastExceptionLocation = lpLocation
        Else
          If mobjLastException.Equals(lpException) = False Then
            mobjLastException = lpException
            mstrLastExceptionLocation = lpLocation
          End If
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace
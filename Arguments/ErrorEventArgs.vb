'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports Documents.Utilities

#End Region

Namespace Arguments

  ''' <summary>
  ''' Contains all the parameters for the ExportError Event
  ''' </summary>
  ''' <remarks></remarks>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public MustInherit Class ErrorEventArgs
    Inherits WriteMessageArgs

#Region "Class Variables"

    Private mobjException As Exception
    Private mobjWorker As BackgroundWorker = Nothing

#End Region

#Region "Public Properties"

    Public ReadOnly Property Exception() As Exception
      Get
        Return mobjException
      End Get
    End Property

    Public ReadOnly Property Worker() As BackgroundWorker
      Get
        Return mobjWorker
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpMessage As String, ByVal lpException As Exception)
      MyBase.New(lpMessage)
      mobjException = lpException
    End Sub

    Public Sub New(ByVal lpMessage As String, ByVal lpException As Exception, ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpMessage)
      mobjException = lpException
      mobjWorker = lpWorker
    End Sub

#End Region

#Region "Private Methods"

    Protected Friend Overrides Function DebuggerIdentifier() As String
      Try
        Return String.Format("{0}; Error={1}:{2}", Message, Exception.GetType.Name, Exception.Message)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
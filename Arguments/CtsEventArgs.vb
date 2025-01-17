'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Text
Imports Documents.Utilities


#End Region


Namespace Arguments

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class CtsEventArgs
    Inherits EventArgs

#Region "Class Variables"

    Private mdatTime As DateTime
    Private mstrEventDescription As String = String.Empty
    Private mblnOperationSucceeded As Boolean = True
    Private mstrMessage As String = String.Empty

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Categorizes the event
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Enables a single event handler for all related events by distinguishing the type or description of the event.</remarks>
    Public Property EventDescription() As String
      Get
        Return mstrEventDescription
      End Get
      Set(ByVal value As String)
        mstrEventDescription = value
      End Set
    End Property

    ''' <summary>
    ''' The time of the event
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Time() As DateTime
      Get
        Return mdatTime
      End Get
      Set(ByVal value As DateTime)
        mdatTime = value
      End Set
    End Property

    ''' <summary>
    ''' Indicates whether or not the document operation succeeded or failed
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property OperationSucceeded() As Boolean
      Get
        Return mblnOperationSucceeded
      End Get
      Set(ByVal value As Boolean)
        mblnOperationSucceeded = value
      End Set
    End Property

    Public Property Message() As String
      Get
        Return mstrMessage
      End Get
      Set(ByVal value As String)
        mstrMessage = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpOperationSucceeded As Boolean)
      Me.New(Now, String.Empty, lpOperationSucceeded, String.Empty)
    End Sub

    Public Sub New(ByVal lpOperationSucceeded As Boolean,
                   ByVal lpMessage As String)
      Me.New(Now, String.Empty, lpOperationSucceeded, lpMessage)
    End Sub

    Public Sub New(ByVal lpEventDescription As String,
                   ByVal lpOperationSucceeded As Boolean,
                   ByVal lpMessage As String)
      Me.New(Now, lpEventDescription, lpOperationSucceeded, lpMessage)
    End Sub

    Public Sub New(ByVal lpTime As DateTime,
                   ByVal lpEventDescription As String,
                   ByVal lpOperationSucceeded As Boolean,
                   ByVal lpMessage As String)
      Try
        Time = lpTime
        EventDescription = lpEventDescription
        OperationSucceeded = lpOperationSucceeded
        Message = lpMessage
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

    Protected Friend Overridable Function DebuggerIdentifier() As String
      Try
        Dim lobjReturnBuilder As New StringBuilder

        lobjReturnBuilder.AppendFormat("Time={0}; Desc={1}; Success={2}; Message={3}",
                                       Time, EventDescription,
                                       OperationSucceeded.ToString, Message)

        Return lobjReturnBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
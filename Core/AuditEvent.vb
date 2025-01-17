' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  AuditEvent.vb
'  Description :  [type_description_here]
'  Created     :  3/13/2012 12:56:48 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Text
Imports Documents.Utilities

#End Region

Namespace Core

  ''' <summary>Used to represent an auditable event of a repository object such as a document or folder.</summary>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class AuditEvent
    Implements INamedItem, IAuditEvent

#Region "Class Variables"

    Private mstrId As String = String.Empty
    Private mdatEventDate As DateTime
    Private mstrName As String = String.Empty
    Private menuStatus As AuditEventStatusEnum
    Private mstrUser As String = String.Empty

#End Region

#Region "Public Properties"

    ''' <summary>Gets or sets the identifier of the audit event.</summary>
    Public Property Id As String Implements IAuditEvent.Id
      Get
        Return mstrId
      End Get
      Set(value As String)
        mstrId = value
      End Set
    End Property

    ''' <summary>Gets or sets the date of the audit event.</summary>
    Public Property EventDate As Date Implements IAuditEvent.EventDate
      Get
        Return mdatEventDate
      End Get
      Set(value As Date)
        mdatEventDate = value
      End Set
    End Property

    ''' <summary>Gets or sets the name of the audit event.</summary>
    Public Property Name As String Implements INamedItem.Name, IAuditEvent.Name
      Get
        Return mstrName
      End Get
      Set(value As String)
        mstrName = value
      End Set
    End Property

    ''' <summary>Gets or sets the status of the audit event.  The status may be Success or Failure.</summary>
    Public Property Status As AuditEventStatusEnum Implements IAuditEvent.Status
      Get
        Return menuStatus
      End Get
      Set(value As AuditEventStatusEnum)
        menuStatus = value
      End Set
    End Property

    ''' <summary>Gets or sets the name of the user that initiated the event.</summary>
    Public Property User As String Implements IAuditEvent.User
      Get
        Return mstrUser
      End Get
      Set(value As String)
        mstrUser = value
      End Set
    End Property

#End Region

#Region "Protected Methods"

    Protected Friend Overridable Function DebuggerIdentifier() As String
      Try
        Dim lobjIdentifierBuilder As New StringBuilder

        lobjIdentifierBuilder.AppendFormat("{0}-{1}:{2} ({3})", Me.Id, Me.User, Me.EventDate, Me.Status.ToString())

        Return lobjIdentifierBuilder.ToString()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
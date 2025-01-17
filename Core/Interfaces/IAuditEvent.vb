' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IAuditEvent.vb
'  Description :  [type_description_here]
'  Created     :  3/13/2012 11:25:09 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"



#End Region

Namespace Core

  Public Interface IAuditEvent

    ''' <summary>Gets or sets the identifier of the audit event.</summary>
    Property Id As String
    ''' <summary>Gets or sets the name of the audit event.</summary>
    Property Name As String
    ''' <summary>Gets or sets the date of the audit event.</summary>
    Property EventDate As DateTime
    ''' <summary>Gets or sets the name of the user that initiated the event.</summary>
    Property User As String
    ''' <summary>Gets or sets the status of the audit event. The status may be Success or Failure.</summary>
    Property Status As AuditEventStatusEnum

  End Interface

  ''' <summary>Used to represent the event status, may be Success or Failure.</summary>
  Public Enum AuditEventStatusEnum
    ''' <summary>The event succeeeded.</summary>
    Success = 0
    ''' <summary>The event failed.</summary>
    Failure = -1
  End Enum

End Namespace
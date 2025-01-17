' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IAuditEvents.vb
'  Description :  [type_description_here]
'  Created     :  3/13/2012 1:02:24 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"



#End Region

Namespace Core

  ''' <summary>Contains a collection of audit events.</summary>
  Public Interface IAuditEvents
    Inherits ICollection(Of IAuditEvent)

    Property Item(id As String) As IAuditEvent
    Overloads Function Contains(id As String) As Boolean

  End Interface

End Namespace
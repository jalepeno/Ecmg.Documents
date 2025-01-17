' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IActionableInformation.vb
'  Description :  Used to describe an operation for the catalog
'  Created     :  11/29/2011 9:37:25 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
'
'  Renamed from IOperationInformation to IActionableInformation and moved to 
'  Ecmg.Cts.Core on 11/2/2013 
'
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"


#End Region

Namespace Core

  Public Interface IActionableInformation

    ReadOnly Property Name As String
    ReadOnly Property DisplayName As String
    ReadOnly Property Description As String
    ReadOnly Property CompanyName As String
    ReadOnly Property ProductName As String
    ReadOnly Property IsExtension As Boolean
    Property ExtensionPath As String

  End Interface

End Namespace
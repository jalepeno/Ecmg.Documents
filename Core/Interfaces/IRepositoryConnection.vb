' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IRepositoryConnection.vb
'  Description :  [type_description_here]
'  Created     :  12/2/2011 8:55:53 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Providers

#End Region

Namespace Core

  Public Interface IRepositoryConnection

    Property Active As Boolean
    Property ConnectionString As String
    ReadOnly Property Properties As IProperties
    ReadOnly Property Provider As IProvider

    ReadOnly Property Repository As Repository

    ReadOnly Property IsDisposed As Boolean

    Property State As ProviderConnectionState
    Property ExportPath As String
    Sub Connect()
    Sub Disconnect()

    Function GetRepository() As Repository

  End Interface

End Namespace
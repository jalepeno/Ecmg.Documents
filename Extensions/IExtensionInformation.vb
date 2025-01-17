' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IExtensionInformation.vb
'  Description :  [type_description_here]
'  Created     :  12/10/2011 6:49:44 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

#End Region

Namespace Extensions

  Public Interface IExtensionInformation

    ReadOnly Property Name As String
    ReadOnly Property DisplayName As String
    ReadOnly Property Description As String
    ReadOnly Property CompanyName As String
    ReadOnly Property ProductName As String
    ReadOnly Property Path As String
    ReadOnly Property IsValid As Boolean
    ReadOnly Property LoadException As Exception
    Sub OnLoadError(ByVal exception As Exception)
    Function ToXmlElementString() As String

  End Interface

End Namespace
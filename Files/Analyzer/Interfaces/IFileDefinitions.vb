'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  IFileDefinitions.vb
'   Description :  [type_description_here]
'   Created     :  1/26/2015 2:11:51 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"


#End Region

Namespace Files
  Public Interface IFileDefinitions
    Inherits IList(Of IFileDefinition)

    Function GetMimeType(lpExtension As String, lpDefaultValue As String) As String

  End Interface

End Namespace
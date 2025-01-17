'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  INameValuePair.vb
'   Description :  [type_description_here]
'   Created     :  12/21/2012 9:19:10 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

Namespace Core

  Public Interface INameValuePair
    Inherits INamedItem

    Property Value As String

  End Interface

End Namespace
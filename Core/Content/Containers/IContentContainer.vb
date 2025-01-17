' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ContentStreamContainer.vb
'  Description :  [type_description_here]
'  Created     :  12/14/2011 8:12:00 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

Namespace Core

  Public Interface IContentContainer

    Property FileName As String
    Property MimeType As String
    Property FileContent As Object
    ReadOnly Property CanRead As Boolean

  End Interface

End Namespace
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  IFileInformation.vb
'   Description :  [type_description_here]
'   Created     :  1/26/2015 9:50:16 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

Namespace Files
  Public Interface IFileInformation

    Property FileType As String
    Property Extension As String
    Property MimeType As String
    Property Popularity As Rating
    Property Remarks As String
    Property RefUrl As String

    ReadOnly Property BasicHash As String

  End Interface

End Namespace
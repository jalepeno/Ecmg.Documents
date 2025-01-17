'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  IFileTestResult.vb
'   Description :  [type_description_here]
'   Created     :  1/28/2015 1:36:08 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

Namespace Files
  Public Interface IFileTestResult

    Property FileType As String
    Property Extension As String
    Property MimeType As String
    Property FilePoints As String
    Property Points As Integer
    Property Percentage As Single
    Property Popularity As Rating
    Function ToString() As String

  End Interface

End Namespace
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  IFileDefinition.vb
'   Description :  [type_description_here]
'   Created     :  1/26/2015 9:43:57 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"


#End Region

Namespace Files
  Public Interface IFileDefinition

    Property FileInfo As IFileInformation

    Property Patterns As IHeaderPatterns

    Property GlobalStrings As IList(Of ByteString)

    Property FrontBlockSize As Integer

    ReadOnly Property InfoHash As String

    Property ItemsFound As Long

  End Interface

End Namespace
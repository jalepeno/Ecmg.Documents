' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  INamedStream.vb
'  Description :  [type_description_here]
'  Created     :  6/11/2012 9:41:55 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.IO

#End Region

Namespace Utilities

  Public Interface INamedStream

    Property FileName As String
    Property Stream As Stream
#If CTSDOTNET = 1 Then
    Function WriteToTempFile() As String
    Sub DeleteTempFile()
    ReadOnly Property TempFilePath As String
#End If

  End Interface

End Namespace
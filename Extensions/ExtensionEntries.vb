' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ExtensionEntries.vb
'  Description :  [type_description_here]
'  Created     :  12/10/2011 7:22:02 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core

#End Region

Namespace Extensions

  Public Class ExtensionEntries
    Inherits CCollection(Of IExtensionInformation)

    'Public Overloads Sub Sort()
    '  Try
    '    MyBase.Sort(New NameComparer)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

  End Class

End Namespace
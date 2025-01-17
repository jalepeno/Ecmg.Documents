' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IVersionIdentifier.vb
'  Description :  Used for identifying a version of a document or object.
'  Created     :  12/19/2011 8:20:16 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

Namespace Core

  Public Interface IVersionIdentifier

    ReadOnly Property MajorVersion As Object
    ReadOnly Property MinorVersion As Object
    ReadOnly Property Version As Object

  End Interface

End Namespace

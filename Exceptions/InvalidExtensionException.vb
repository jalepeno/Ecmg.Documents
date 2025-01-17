' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  InvalidExtensionException.vb
'  Description :  [type_description_here]
'  Created     :  11/16/2011 1:17:53 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

Namespace Exceptions

  Public Class InvalidExtensionException
    Inherits CtsException

    Public Sub New(ByVal Message As String)
      MyBase.New(Message)
    End Sub

    Public Sub New(ByVal Message As String,
                   ByVal InnerException As Exception)
      MyBase.New(Message, InnerException)
    End Sub

  End Class

End Namespace
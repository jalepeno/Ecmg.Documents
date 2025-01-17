'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  InvalidTransformationTargetException.vb
'   Description :  [type_description_here]
'   Created     :  3/5/2015 3:33:57 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

Namespace Exceptions

  Public Class InvalidTransformationTargetException
    Inherits CtsException

    Public Sub New()
      MyBase.New("Unexpected target type")
    End Sub

    Public Sub New(ByVal Message As String)
      MyBase.New(Message)
    End Sub

    Public Sub New(ByVal Message As String, ByVal InnerException As Exception)
      MyBase.New(Message, InnerException)
    End Sub

  End Class

End Namespace
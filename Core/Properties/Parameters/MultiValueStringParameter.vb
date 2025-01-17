'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  MultiValueStringParameter.vb
'   Description :  [type_description_here]
'   Created     :  6/18/2013 2:09:47 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

Namespace Core

  Public Class MultiValueStringParameter
    Inherits MultiValueParameter

#Region "Constructors"

    Public Sub New()
      MyBase.New(PropertyType.ecmString)
    End Sub

    Protected Friend Sub New(ByVal lpName As String)
      MyBase.New(PropertyType.ecmString, lpName)
    End Sub

    Protected Friend Sub New(ByVal lpName As String, ByVal lpSystemName As String, ByVal lpValues As Values)
      MyBase.New(PropertyType.ecmString, lpName, lpSystemName, lpValues)
    End Sub

#End Region

  End Class

End Namespace
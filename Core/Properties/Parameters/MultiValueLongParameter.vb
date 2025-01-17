'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  MultiValueLongParameter.vb
'   Description :  [type_description_here]
'   Created     :  6/18/2013 2:50:11 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

Namespace Core

  Public Class MultiValueLongParameter
    Inherits MultiValueParameter

#Region "Constructors"

    Public Sub New()
      MyBase.New(PropertyType.ecmLong)
    End Sub

    Protected Friend Sub New(ByVal lpName As String)
      MyBase.New(PropertyType.ecmLong, lpName)
    End Sub

    Protected Friend Sub New(ByVal lpName As String, ByVal lpSystemName As String, ByVal lpValues As Values)
      MyBase.New(PropertyType.ecmLong, lpName, lpSystemName, lpValues)
    End Sub

#End Region

  End Class

End Namespace
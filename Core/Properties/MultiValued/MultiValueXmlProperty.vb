'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"



#End Region

Namespace Core

  Public Class MultiValueXmlProperty
    Inherits MultiValueProperty

#Region "Constructors"

    Public Sub New()
      MyBase.New(PropertyType.ecmXml)
    End Sub

    Public Sub New(ByVal lpName As String)
      MyBase.New(PropertyType.ecmXml, lpName)
    End Sub

    Public Sub New(ByVal lpName As String, ByVal lpSystemName As String, ByVal lpValues As Values)
      MyBase.New(PropertyType.ecmXml, lpName, lpSystemName, lpValues)
    End Sub

#End Region

  End Class

End Namespace
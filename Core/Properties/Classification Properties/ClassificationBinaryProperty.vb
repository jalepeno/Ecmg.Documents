'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"



#End Region

Namespace Core

  Public Class ClassificationBinaryProperty
    Inherits ClassificationProperty

#Region "Constructors"

    Friend Sub New()
      MyBase.New(PropertyType.ecmBinary)
    End Sub

    Friend Sub New(ByVal lpName As String)
      MyBase.New(PropertyType.ecmBinary, lpName)
    End Sub

    Friend Sub New(ByVal lpName As String, ByVal lpSystemName As String)
      MyBase.New(PropertyType.ecmBinary, lpName, lpSystemName)
    End Sub

#End Region

  End Class

End Namespace
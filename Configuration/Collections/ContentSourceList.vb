'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core

#End Region

Namespace Configuration

  Public Class ContentSourceList
    Inherits NamedList

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpName As String)
      MyBase.New(lpName)
    End Sub

#End Region

  End Class

End Namespace

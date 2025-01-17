'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"



#End Region

Namespace Transformations

  <Serializable()>
  Public Class PromoteMetadataAction
    Inherits CreatePropertyAction

#Region "Class Constants"

    Private Const ACTION_NAME As String = "PromoteMetadata"

#End Region

#Region "Class Variables"

    Private mstrMetadataTagName As String = String.Empty

#End Region

#Region "Public Properties"

    Public Overrides ReadOnly Property ActionName As String
      Get
        Return ACTION_NAME
      End Get
    End Property

#End Region

#Region "Constructors"

#End Region

  End Class

End Namespace

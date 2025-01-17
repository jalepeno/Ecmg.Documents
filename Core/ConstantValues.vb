' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  Constants.vb
'  Description :  Contains shared constant values
'  Created     :  11/29/2011 10:12:29 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

Namespace Core

  Public Class ConstantValues

#Region "Constant Values"

    Private Const COMPANY_NAME As String = "Conteage Corp"
    Private Const PRODUCT_NAME As String = "Content Transformation Services"

#End Region

#Region "Constructors"

    Private Sub New()
      ' We do not want instances created for this class
    End Sub

#End Region

#Region "Public Shared Properties"

    Public Shared ReadOnly Property CompanyName As String
      Get
        Return COMPANY_NAME
      End Get
    End Property

    Public Shared ReadOnly Property ProductName As String
      Get
        Return PRODUCT_NAME
      End Get
    End Property

#End Region

  End Class

End Namespace
' ********************************************************************************
' '  Document    :  UrlFactory.vb
' '  Description :  [type_description_here]
' '  Created     :  11/21/2012-1:45:53
' '  <copyright company="ECMG">
' '      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
' '      Copying or reuse without permission is strictly forbidden.
' '  </copyright>
' ********************************************************************************

#Region "Imports"

#End Region

Namespace Core

  Public Class UrlFactory
    Inherits StringFactory

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(lpFormatString As String, ParamArray lpParameters() As String)
      MyBase.New(lpFormatString, lpParameters)
    End Sub

#End Region

  End Class

End Namespace
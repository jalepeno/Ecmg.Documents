'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Namespace Scripting

  Public Interface IMethod

#Region "Public Properties"

    Property Name() As String
    Property Parameters() As Core.ECMProperties

#End Region

#Region "Public Methods"

    'Function Execute() As MethodResult
    'Function LastResult() As MethodResult

#End Region

  End Interface

End Namespace
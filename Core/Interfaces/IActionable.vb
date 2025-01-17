' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IActionable.vb
'  Description :  [type_description_here]
'  Created     :  11/2/2013 3:17:42 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"


#End Region

Namespace Core

  Public Interface IActionable
    Inherits IActionableBase

    Function Execute() As Transformations.ActionResult

    ''' <summary>
    ''' Called to execute.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function Execute(ByRef lpErrorMessage As String) As Transformations.ActionResult

    ''' <summary>
    ''' Used to initialize internal properties from parameters.
    ''' </summary>
    ''' <remarks></remarks>
    Sub InitializeParameterValues()

    ReadOnly Property ResultDetail As IActionableResult

  End Interface

End Namespace
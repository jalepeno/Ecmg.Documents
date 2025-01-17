'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  IFileTestResults.vb
'   Description :  [type_description_here]
'   Created     :  1/28/2015 1:36:16 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"


#End Region

Namespace Files
  Public Interface IFileTestResults
    Inherits IList(Of IFileTestResult)

    ReadOnly Property PrimaryResult As IFileTestResult

  End Interface

End Namespace
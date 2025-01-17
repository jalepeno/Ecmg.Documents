'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  ITransformationConfiguration.vb
'   Description :  [type_description_here]
'   Created     :  4/24/2014 2:54:17 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
Imports Documents.Transformations

Namespace Core

  Public Interface ITransformationConfiguration

    Property Name As String

    ReadOnly Property DisplayName As String

    Property Description As String

    Property Actions As IActionItems

    Function ToTransformation() As Transformation

    Function ToXmlString() As String

  End Interface

End Namespace
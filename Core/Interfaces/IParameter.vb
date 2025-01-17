' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IParameter.vb
'  Description :  [type_description_here]
'  Created     :  11/18/2011 3:23:13 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"


Imports System.Xml

#End Region

Namespace Core

  Public Interface IParameter
    Inherits IProperty
    Inherits ICloneable

    Function ToXmlString() As String
    Sub InitializeFromXmlNode(lpXmlNode As XmlNode)

  End Interface

End Namespace
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IActionableResult.vb
'  Description :  [type_description_here]
'  Created     :  11/2/2013 1:06:40 9M
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

Namespace Core

  Public Interface IActionableResult
    Inherits IDisposable

    Property Name As String
    Property Result As Result
    Property ProcessedMessage As String

    Property StartTime As DateTime
    Property FinishTime As DateTime
    Property TotalProcessingTime As TimeSpan

    ReadOnly Property Parent As IActionableBase

    Function ToJsonString() As String
    Function ToXmlString() As String
    Function ToXmlElementString() As String

  End Interface

End Namespace

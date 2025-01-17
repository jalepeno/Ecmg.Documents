' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IJsonSerializable.vb
'  Description :  [type_description_here]
'  Created     :  01/22/2024 10:41:13 PM
'  <copyright company="Conteage">
'      Copyright (c) Conteage Corp. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

Namespace Core

  Public Interface IJsonSerializable(Of T)

    Function ToJson() As String

    Function FromJson(lpJson As String) As T

  End Interface

End Namespace
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ServerThrottlingException.vb
'  Description :  [type_description_here]
'  Created     :  9/25/2023 12:07:44 PM
'  <copyright company="Conteage">
'      Copyright (c) Conteage Corp All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"


#End Region

Namespace Exceptions

  Public Class ServerThrottlingException
    Inherits CtsException

#Region "Constants"

    Private Const MESSAGE As String = "The server is throttling the client."

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub
    Public Sub New(ByVal message As String)
      MyBase.New(message)
    End Sub

    Public Sub New(ByVal inner As Exception)
      MyBase.New(MESSAGE, inner)
    End Sub

    Public Sub New(ByVal message As String, ByVal inner As Exception)
      MyBase.New(message, inner)
    End Sub

#End Region

  End Class

End Namespace
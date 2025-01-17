' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  DocumentHasNoContentException.vb
'  Description :  [type_description_here]
'  Created     :  12/8/2011 2:43:42 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core

#End Region

''' <summary>
''' Contains the defined exceptions of the Cts framework
''' </summary>
Namespace Exceptions

  ''' <summary>
  ''' Exception thrown as a result of a document 
  ''' having no content when content is expected.
  ''' </summary>
  ''' <remarks></remarks>
  Public Class DocumentHasNoContentException
    Inherits DocumentException

#Region "Constructors"

    Public Sub New(ByVal document As Document, ByVal message As String)
      MyBase.New(document, message)
    End Sub

#End Region

  End Class

End Namespace
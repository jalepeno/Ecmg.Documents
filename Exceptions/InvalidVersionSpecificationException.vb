'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Core

Namespace Exceptions
  ''' <summary>Thrown when a version is specified improperly.</summary>
  ''' <remarks>Typically occurs in conjunction with document export operations.</remarks>
  Public Class InvalidVersionSpecificationException
    Inherits DocumentException

#Region "Constructors"

    ''' <summary>
    ''' Creates a new DocumentException with the specified document id.
    ''' </summary>
    ''' <param name="id">The Document Id</param>
    ''' <remarks>Initializes the VersionId to zero and the document object to a null object reference</remarks>
    Public Sub New(ByVal id As String, ByVal message As String)
      MyBase.New(id, "0", message)
    End Sub

    ''' <summary>
    ''' Creates a new DocumentException with the specified document.
    ''' </summary>
    ''' <param name="document">The Document object</param>
    ''' <remarks>Initializes the VersionId to zero</remarks>
    Public Sub New(ByVal document As Document, ByVal message As String)
      MyBase.New(document, "0", message)
    End Sub

#End Region

  End Class

End Namespace

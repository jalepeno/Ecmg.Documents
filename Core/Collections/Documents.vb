'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Core
  ''' <summary>Collection of Document objects.</summary>
  Public Class Documents
    Inherits CCollection(Of Document)

#Region "Class Variables"

#End Region

#Region "Public Properties"

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Checks for the existence of the document in the collection based on the id, then the name.
    ''' </summary>
    ''' <param name="name">The Id or Name of the document.</param>
    ''' <returns>True if found, otherwise false.</returns>
    ''' <remarks></remarks>
    Public Overloads Function Contains(name As String) As Boolean
      Try
        Return Contains(name, Nothing)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Checks for the existence of the document in the collection based on the id, then the name.
    ''' </summary>
    ''' <param name="name">The Id or Name of the document.</param>
    ''' <param name="lpFoundDocument">A return reference to the found document, if available.</param>
    ''' <returns>True if found, otherwise false.</returns>
    ''' <remarks></remarks>
    Public Overloads Function Contains(name As String, ByRef lpFoundDocument As Document) As Boolean
      Try

        lpFoundDocument = MyBase.FirstOrDefault(Function(document) document.ID = name)

        If lpFoundDocument IsNot Nothing Then
          Return True
        End If

        lpFoundDocument = MyBase.FirstOrDefault(Function(document) document.Name = name)

        If lpFoundDocument IsNot Nothing Then
          Return True
        End If

        Return False

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Providers
Imports Documents.Utilities

#End Region

Namespace Migrations

  ''' <summary>
  ''' Used for migrating documents or folders from a source repository to a destination
  ''' repository.
  ''' </summary>
  Public Class Migrator

#Region "Public Methods"
    Public Shared Sub ValidateProperties(ByVal lpDocument As Document, ByVal lpImporter As Providers.IDocumentImporter)

      Try

        ' Validate the properties against the destination document class
        If CType(lpImporter, Providers.CProvider).SupportsInterface(Providers.ProviderClass.Classification) Then
          Dim lstrValidationError As String = String.Empty

          Dim lobjInvalidProperties As InvalidProperties

          ' First check to see if the class is valid
          'If CType(Importer, Providers.IClassification).DocumentClasses.ClassExists(lpDocument.DocumentClass) Then
          If CType(lpImporter, Providers.IClassification).DocumentClass(lpDocument.DocumentClass) IsNot Nothing Then

            ' Then check to see if the properties are valid
            lobjInvalidProperties = CType(lpImporter, CProvider).
              FindInvalidProperties(lpDocument, PropertyScope.BothDocumentAndVersionProperties)

            'If lobjInvalidProperties.Count > 0 Then
            '  ' Do not import the document
            '  lstrValidationError = String.Format("Document '{0}' contains properties invalid for importing to class '{1}': '{2}'", _
            '                                                 lpDocument.ID, _
            '                                                 lpDocument.DocumentClass, _
            '                                                 lobjInvalidProperties.ListProperties)

            '  Throw New DocumentValidationException(lpDocument.ID, lstrValidationError)

            'End If

          Else
            ' The class does not even exist in the destination
            lstrValidationError = String.Format("The document class '{0}' is not defined for the target {1} repository '{2}'.",
                                                lpDocument.DocumentClass,
                                                CType(lpImporter, CProvider).Information.ProductName,
                                                CType(lpImporter, CProvider).ContentSource.Name)

            Throw New DocumentValidationException(lpDocument.ID, lstrValidationError)

          End If

        End If

      Catch ValEx As DocumentValidationException
        ' Just throw it back
        Throw
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub
#End Region

  End Class

End Namespace
'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Data
Imports Documents.Arguments
Imports Documents.Core
Imports Documents.Data
Imports Documents.Security

#End Region

''' <summary>
''' Contains objects related to repository providers
''' </summary>
Namespace Providers

#Region "Provider Objects"

#Region "Event Delegates"

  ''' <summary>Delegate event handler for the DocumentExported event.</summary>
  Public Delegate Sub DocumentExportEventHandler(ByVal sender As Object, ByRef e As DocumentExportedEventArgs)
  ''' <summary>Delegate event handler for the FolderDocumentExported event.</summary>
  Public Delegate Sub FolderDocumentExportEventHandler(ByVal sender As Object, ByRef e As FolderDocumentExportedEventArgs)
  ''' <summary>Delegate event handler for the FolderExported event.</summary>
  Public Delegate Sub FolderExportedEventHandler(ByVal sender As Object, ByRef e As FolderExportedEventArgs)
  ''' <summary>Delegate event handler for the DocumentImported event.</summary>
  Public Delegate Sub DocumentImportedEventHandler(ByVal sender As Object, ByVal e As DocumentImportedEventArgs)
  ''' <summary>Delegate event handler for the DocumentExportError event.</summary>
  Public Delegate Sub DocumentExportErrorEventHandler(ByVal sender As Object, ByVal e As DocumentExportErrorEventArgs)
  ''' <summary>Delegate event handler for the DocumentImportError event.</summary>
  Public Delegate Sub DocumentImportErrorEventHandler(ByVal sender As Object, ByVal e As DocumentImportErrorEventArgs)

  ''' <summary>Delegate event handler for the DocumentExportMessage event.</summary>
  Public Delegate Sub DocumentExportMessageEventHandler(ByVal sender As Object, ByVal e As WriteMessageArgs)

  ''' <summary>Delegate event handler for the DocumentImportMessage event.</summary>
  Public Delegate Sub DocumentImportMessageEventHandler(ByVal sender As Object, ByVal e As WriteMessageArgs)

  ''' <summary>Delegate folder event handler.</summary>
  Public Delegate Sub FolderEventHandler(ByVal sender As Object, ByVal e As FolderEventArgs)

  ''' <summary>Delegate event handler for the FolderExportError event.</summary>
  Public Delegate Sub FolderExportErrorEventHandler(ByVal sender As Object, ByVal e As FolderExportErrorEventArgs)

  ''' <summary>Delegate event handler for the ObjectExportError event.</summary>
  Public Delegate Sub ObjectExportErrorEventHandler(ByVal sender As Object, ByVal e As ObjectExportErrorEventArgs)

  ''' <summary>Delegate event handler for Searching</summary>
  Public Delegate Sub SearchUpdateEventHandler(ByVal sender As Object, ByVal e As SearchEventArgs)
  ''' <summary>Delegate event handler for Searching</summary>
  Public Delegate Sub SearchCompleteEventHandler(ByVal sender As Object, ByVal e As SearchEventArgs)

  ''' <summary>Delegate event handler for the DocumentClassImported event.</summary>
  Public Delegate Sub DocumentClassImportedEventHandler(ByVal sender As Object, ByVal e As DocumentClassImportedEventArgs)

  ''' <summary>Delegate event handler for the ConnectionStateChanged event.</summary>
  Public Delegate Sub ConnectionStateChangedEventHandler(ByVal sender As Object, ByRef e As ConnectionStateChangedEventArgs)

  ''' <summary>Delegate event handler for the ProviderPropertyValueChanged event.</summary>
  Public Delegate Sub ProviderPropertyValueChangedEventHandler(ByVal sender As Object, ByRef e As ProviderPropertyValueChangedEventArgs)

  ''' <summary>Delegate event handler for the DocumentExported event.</summary>
  Public Delegate Sub AnnotationExportEventHandler(ByVal sender As Object, ByRef e As AnnotationExportedEventArgs)

  ''' <summary>Delegate event handler for the DocumentExportMessage event.</summary>
  Public Delegate Sub AnnotationExportMessageEventHandler(ByVal sender As Object, ByVal e As WriteMessageArgs)


#Region "IBasicContentServer Event Delegates"

  Public Delegate Sub DocumentAddedEventHandler(ByVal sender As Object, ByVal e As DocumentAddedEventArgs)
  Public Delegate Sub DocumentCopiedOutEventHandler(ByVal sender As Object, ByVal e As DocumentCopiedOutEventArgs)
  Public Delegate Sub DocumentCheckedOutEventHandler(ByVal sender As Object, ByVal e As DocumentCheckedOutEventArgs)
  Public Delegate Sub DocumentCheckedInEventHandler(ByVal sender As Object, ByVal e As DocumentCheckedInEventArgs)
  Public Delegate Sub DocumentCheckOutCancelledEventHandler(ByVal sender As Object, ByVal e As DocumentCheckoutCancelledEventArgs)
  Public Delegate Sub DocumentDeletedEventHandler(ByVal sender As Object, ByVal e As DocumentDeletedEventArgs)
  Public Delegate Sub DocumentUpdatedEventHandler(ByVal sender As Object, ByVal e As DocumentUpdatedEventArgs)
  Public Delegate Sub DocumentFiledEventHandler(ByVal sender As Object, ByVal e As DocumentFiledEventArgs)
  Public Delegate Sub DocumentUnFiledEventHandler(ByVal sender As Object, ByVal e As DocumentUnFiledEventArgs)
  Public Delegate Sub DocumentEventHandler(ByVal sender As Object, ByVal e As DocumentEventArgs)

#End Region

#End Region

#Region "Provider Interfaces"

  ''' <summary>
  ''' Interface used for objects needing to Authorize.
  ''' </summary>
  ''' <remarks></remarks>
  Public Interface IAuthorization
    Function Login(ByVal lpUserName As String, ByVal lpPassword As String) As Boolean
    Function Logout() As Boolean
  End Interface

  ''' <summary>
  ''' Interface to be implemented by all providers that provide basic content services
  ''' </summary>
  ''' <remarks></remarks>
  Public Interface IBasicContentServicesProvider

#Region "Events"

    ''' <summary>
    ''' Raised upon completion of AddDocument method
    ''' </summary>
    ''' <remarks>Note: May be raised on both success and failure, check the OperationSucceeded property of 'e' to determine status</remarks>
    Event DocumentAdded As DocumentAddedEventHandler

    ''' <summary>
    ''' Raised upon error exporting document.
    ''' </summary>
    ''' <remarks></remarks>
    Event DocumentExportError As DocumentExportErrorEventHandler

    ''' <summary>
    ''' Raised upon error adding document.
    ''' </summary>
    ''' <remarks></remarks>
    Event DocumentImportError As DocumentImportErrorEventHandler

    ''' <summary>
    ''' Raised upon completion of CopyOutDocument method
    ''' </summary>
    ''' <remarks>Note: May be raised on both success and failure, check the OperationSucceeded property of 'e' to determine status</remarks>
    Event DocumentCopiedOut As DocumentCopiedOutEventHandler

    ''' <summary>
    ''' Raised upon completion of CheckoutDocument method
    ''' </summary>
    ''' <remarks>Note: May be raised on both success and failure, check the OperationSucceeded property of 'e' to determine status</remarks>
    Event DocumentCheckedOut As DocumentCheckedOutEventHandler

    ''' <summary>
    ''' Raised upon completion of CheckinDocument method
    ''' </summary>
    ''' <remarks>Note: May be raised on both success and failure, check the OperationSucceeded property of 'e' to determine status</remarks>
    Event DocumentCheckedIn As DocumentCheckedInEventHandler

    ''' <summary>
    ''' Raised upon completion of CancelCheckoutDocument method
    ''' </summary>
    ''' <remarks>Note: May be raised on both success and failure, check the OperationSucceeded property of 'e' to determine status</remarks>
    Event DocumentCheckOutCancelled As DocumentCheckOutCancelledEventHandler

    ''' <summary>
    ''' Raised upon completion of DeleteDocument method
    ''' </summary>
    ''' <remarks>Note: May be raised on both success and failure, check the OperationSucceeded property of 'e' to determine status</remarks>
    Event DocumentDeleted As DocumentDeletedEventHandler

    ''' <summary>
    ''' Raised upon completion of UpdateDocument method
    ''' </summary>
    ''' <remarks>Note: May be raised on both success and failure, check the OperationSucceeded property of 'e' to determine status</remarks>
    Event DocumentUpdated As DocumentUpdatedEventHandler

    ''' <summary>
    ''' Raised upon completion of FileDocument method
    ''' </summary>
    ''' <remarks>Note: May be raised on both success and failure, check the OperationSucceeded property of 'e' to determine status</remarks>
    Event DocumentFiled As DocumentFiledEventHandler

    ''' <summary>
    ''' Raised upon completion of UnFileDocument method
    ''' </summary>
    ''' <remarks>Note: May be raised on both success and failure, check the OperationSucceeded property of 'e' to determine status</remarks>
    Event DocumentUnFiled As DocumentUnFiledEventHandler

    ''' <summary>
    ''' Raised upon completion of all document methods
    ''' </summary>
    ''' <remarks>Note: May be raised on both success and failure, check the OperationSucceeded property of 'e' to determine status.  Check the value of EventDescription to determine which method raised the event.</remarks>
    Event DocumentEvent As DocumentEventHandler

#End Region

#Region "Methods"

    ''' <summary>
    ''' Gets the document metadata as a Document object from the repository
    ''' </summary>
    ''' <param name="lpId">The document identifier</param>
    ''' <returns>A Cts.Core.Document object reference</returns>
    ''' <remarks>This method always returns a document object with no Content values.  
    ''' If the content is required in addition to the metadata it is recommended to 
    ''' use the GetDocumentWithContent method.</remarks>
    Function GetDocumentWithoutContent(ByVal lpId As String) As Document

    ''' <summary>
    ''' Gets the document metadata as a Document object from the repository
    ''' </summary>
    ''' <param name="lpId">The document identifier</param>
    ''' <param name="lpPropertyFilter">
    ''' A list of properties to return for the document.  
    ''' The properties returned will be limited to those in the list.
    ''' </param>
    ''' <returns>A Cts.Core.Document object reference</returns>
    ''' <remarks>This method always returns a document object with no Content values.  
    ''' If the content is required in addition to the metadata it is recommended to 
    ''' use the GetDocumentWithContent method.</remarks>
    Function GetDocumentWithoutContent(ByVal lpId As String, ByVal lpPropertyFilter As List(Of String)) As Document

    ''' <summary>
    ''' Gets the document and content as a Document object from the repository
    ''' </summary>
    ''' <param name="lpId">The document identifier</param>
    ''' <param name="lpDestinationFolder">The fully qualified path to the folder to copy the content file(s) to</param>
    ''' <returns>A Cts.Core.Document object reference</returns>
    ''' <remarks>This method always returns a document object with Content values.  
    ''' If the content is not required in addition to the metadata please it is recommended 
    ''' to use the GetDocumentWithoutContent method.</remarks>
    Function GetDocumentWithContent(ByVal lpId As String, ByVal lpDestinationFolder As String) As Document

    ''' <summary>
    ''' Gets the document and content as a Document object from the repository
    ''' </summary>
    ''' <param name="lpId">The document identifier</param>
    ''' <param name="lpDestinationFolder">The fully qualified path to the folder to copy the content file(s) to</param>
    ''' <param name="lpStorageType">The document storage type</param>
    ''' <returns>A Cts.Core.Document object reference</returns>
    ''' <remarks>This method always returns a document object with Content values.  
    ''' If the content is not required in addition to the metadata please it is recommended 
    ''' to use the GetDocumentWithoutContent method.</remarks>
    Function GetDocumentWithContent(ByVal lpId As String, ByVal lpDestinationFolder As String, ByVal lpStorageType As Content.StorageTypeEnum) As Document

    ''' <summary>
    ''' Adds a new document to the repository
    ''' </summary>
    ''' <param name="lpDocument">The Cts document object to load</param>
    ''' ''' <param name="lpAsMajorVersion">Determines whether or not the document will be checked in as a major or minor version</param>
    ''' <returns>True if successful, false if failed</returns>
    ''' <remarks>If the caller does not have sufficient rights to add the document, an InsufficientRightsException will be thrown</remarks>
    Function AddDocument(ByVal lpDocument As Document, ByVal lpAsMajorVersion As Boolean) As Boolean

    ''' <summary>
    ''' Adds a new document to the repository
    ''' </summary>
    ''' <param name="lpDocument">The Cts document object to load</param>
    ''' <returns>True if successful, false if failed</returns>
    ''' <remarks>If the caller does not have sufficient rights to add the document, an InsufficientRightsException will be thrown</remarks>
    Function AddDocument(ByVal lpDocument As Document) As Boolean

    ''' <summary>
    ''' Adds a new document to the repository and files it in the specified folder path
    ''' </summary>
    ''' <param name="lpDocument">The Cts document object to load</param>
    ''' <param name="lpFolderPath">The repository folder to file the new document in</param>
    ''' <returns>True if successful, false if failed</returns>
    ''' <remarks>If the caller does not have sufficient rights to add the document, an InsufficientRightsException will be thrown</remarks>
    Function AddDocument(ByVal lpDocument As Document, ByVal lpFolderPath As String) As Boolean

    ''' <summary>
    ''' Adds a new document to the repository and files it in the specified folder path
    ''' </summary>
    ''' <param name="lpDocument">The Cts document object to load</param>
    ''' <param name="lpFolderPath">The repository folder to file the new document in</param>
    ''' <param name="lpAsMajorVersion">Determines whether or not the document will be checked in as a major or minor version</param>
    ''' <returns>True if successful, false if failed</returns>
    ''' <remarks>If the caller does not have sufficient rights to add the document, an InsufficientRightsException will be thrown</remarks>
    Function AddDocument(ByVal lpDocument As Document, ByVal lpFolderPath As String, ByVal lpAsMajorVersion As Boolean) As Boolean

    ''' <summary>
    ''' Copies out the content of the latest document to the specified destination folder
    ''' </summary>
    ''' <param name="lpId">The document identifier</param>
    ''' <param name="lpDestinationFolder">The fully qualified path to the folder to copy the content file(s) to</param>
    ''' <param name="lpOutputFileNames">A string array returned by reference with the fully qualified file names of the content elements copied out</param>
    ''' <returns>True if successful, false if failed</returns>
    ''' <remarks>The destination folder should exist before calling this method.</remarks>
    Function CopyOutDocument(ByVal lpId As String, ByVal lpDestinationFolder As String, ByRef lpOutputFileNames As String()) As Boolean

    ''' <summary>
    ''' Returns the Contents of the latest document version
    ''' </summary>
    ''' <param name="lpId">The document identifier</param>
    ''' <param name="lpVersionScope">The version scope to target</param>
    ''' <param name="lpMaxVersionCount">The maximum number of versions to target.  Only applies to FirstNVersions and LastNVersions.</param>
    ''' <returns>Contents collection</returns>
    ''' <remarks></remarks>
    Function GetDocumentContents(lpId As String, lpVersionScope As VersionScopeEnum, lpMaxVersionCount As Integer) As Contents

    ''' <summary>
    ''' Returns the Contents of the latest document version
    ''' </summary>
    ''' <param name="lpId">The document identifier</param>
    ''' <returns>Contents collection</returns>
    ''' <remarks></remarks>
    Function GetDocumentContents(lpId As String) As Contents

    ''' <summary>
    ''' Checks out the document from the repository
    ''' </summary>
    ''' <param name="lpId">The document identifier</param>
    ''' <param name="lpDestinationFolder">The local or network folder to check the document out to</param>
    ''' <param name="lpOutputFileNames">A string array to capture the path(s) of the content element(s) checked out from the repository.</param>
    ''' <returns>True if successful, false if failed</returns>
    ''' <remarks>If the document is already checked out, a DocumentAlreadyOutException will be thrown.  
    ''' If the caller does not have sufficient rights to checkout the document, an InsufficientRightsException will be thrown.</remarks>
    Function CheckoutDocument(ByVal lpId As String, ByVal lpDestinationFolder As String, ByRef lpOutputFileNames As String()) As Boolean

    ''' <summary>
    ''' Cancels the checkout of a document
    ''' </summary>
    ''' <param name="lpId">The document identifier</param>
    ''' <returns>True if successful, false if failed</returns>
    ''' <remarks>If the document is not checked out, a DocumentNotCheckedOutException will be thrown.  
    ''' If the caller does not have sufficient rights to checkout the document, an InsufficientRightsException will be thrown.</remarks>
    Function CancelCheckoutDocument(ByVal lpId As String) As Boolean

    ''' <summary>
    ''' Checks a document back into the repository
    ''' </summary>
    ''' <param name="lpId">The document identifier</param>
    ''' <param name="lpContentPath">The fully qualifed path to the file to checkin</param>
    ''' <param name="lpAsMajorVersion">Determines whether or not the document will be checked in as a major or minor version</param>
    ''' <returns>True if successful, false if failed</returns>
    ''' <remarks>If the document is not checked out, a DocumentNotCheckedOutException will be thrown.  
    ''' If the caller does not have sufficient rights to checkout the document, an InsufficientRightsException will be thrown.</remarks>
    Function CheckinDocument(ByVal lpId As String,
                             ByVal lpContentPath As String, ByVal lpAsMajorVersion As Boolean) As Boolean

    ''' <summary>
    ''' Checks a document back into the repository
    ''' </summary>
    ''' <param name="lpId">The document identifier</param>
    ''' <param name="lpContentContainer">A container for the file to checkin</param>
    ''' <param name="lpAsMajorVersion">Determines whether or not the document will be checked in as a major or minor version</param>
    ''' <returns>True if successful, false if failed</returns>
    ''' <remarks>If the document is not checked out, a DocumentNotCheckedOutException will be thrown.  
    ''' If the caller does not have sufficient rights to checkout the document, an InsufficientRightsException will be thrown.</remarks>
    Function CheckinDocument(ByVal lpId As String,
                             ByVal lpContentContainer As IContentContainer, ByVal lpAsMajorVersion As Boolean) As Boolean

    ''' <summary>
    ''' Checks a document back into the repository
    ''' </summary>
    ''' <param name="lpId">The document identifier</param>
    ''' <param name="lpContentPaths">An array of strings containing the fully qualified paths of all of the content elements to checkin</param>
    ''' <param name="lpAsMajorVersion">Determines whether or not the document will be checked in as a major or minor version</param>
    ''' <returns>True if successful, false if failed</returns>
    ''' <remarks>If the document is not checked out, a DocumentNotCheckedOutException will be thrown.  
    ''' If the caller does not have sufficient rights to checkout the document, an InsufficientRightsException will be thrown.</remarks>
    Function CheckinDocument(ByVal lpId As String,
                             ByVal lpContentPaths As String(),
                             ByVal lpAsMajorVersion As Boolean) As Boolean

    ' <Added by: Ernie at: 3/26/2012-8:32:05 AM on machine: ERNIE-M4400>
    ' Added signatures to allow the caller to provide property values to set on the new version.

    ''' <summary>
    ''' Checks a document back into the repository
    ''' </summary>
    ''' <param name="lpId">The document identifier</param>
    ''' <param name="lpContentPath">The fully qualifed path to the file to checkin</param>
    ''' <param name="lpAsMajorVersion">Determines whether or not the document will be checked in as a major or minor version</param>
    ''' <param name="lpProperties">A collection of properties to set on the new version</param>
    ''' <returns>True if successful, false if failed</returns>
    ''' <remarks>If the document is not checked out, a DocumentNotCheckedOutException will be thrown.  
    ''' If the caller does not have sufficient rights to checkout the document, an InsufficientRightsException will be thrown.</remarks>
    Function CheckinDocument(ByVal lpId As String,
                             ByVal lpContentPath As String,
                             ByVal lpAsMajorVersion As Boolean,
                             ByVal lpProperties As IProperties) As Boolean

    ''' <summary>
    ''' Checks a document back into the repository
    ''' </summary>
    ''' <param name="lpId">The document identifier</param>
    ''' <param name="lpContentContainer">A container for the file to checkin</param>
    ''' <param name="lpAsMajorVersion">Determines whether or not the document will be checked in as a major or minor version</param>
    ''' <param name="lpProperties">A collection of properties to set on the new version</param>
    ''' <returns>True if successful, false if failed</returns>
    ''' <remarks>If the document is not checked out, a DocumentNotCheckedOutException will be thrown.  
    ''' If the caller does not have sufficient rights to checkout the document, an InsufficientRightsException will be thrown.</remarks>
    Function CheckinDocument(ByVal lpId As String,
                            ByVal lpContentContainer As IContentContainer,
                            ByVal lpAsMajorVersion As Boolean,
                            ByVal lpProperties As IProperties) As Boolean

    ''' <summary>
    ''' Checks a document back into the repository
    ''' </summary>
    ''' <param name="lpId">The document identifier</param>
    ''' <param name="lpContentPaths">An array of strings containing the fully qualified paths of all of the content elements to checkin</param>
    ''' <param name="lpAsMajorVersion">Determines whether or not the document will be checked in as a major or minor version</param>
    ''' <param name="lpProperties">A collection of properties to set on the new version</param>
    ''' <returns>True if successful, false if failed</returns>
    ''' <remarks>If the document is not checked out, a DocumentNotCheckedOutException will be thrown.  
    ''' If the caller does not have sufficient rights to checkout the document, an InsufficientRightsException will be thrown.</remarks>
    Function CheckinDocument(ByVal lpId As String,
                            ByVal lpContentPaths As String(),
                            ByVal lpAsMajorVersion As Boolean,
                            ByVal lpProperties As IProperties) As Boolean

    ' </Added by: Ernie at: 3/26/2012-8:32:05 AM on machine: ERNIE-M4400>

    ''' <summary>
    ''' Checks to see if a document is currently checked out.
    ''' </summary>
    ''' <param name="lpId">The document identifier</param>
    ''' <returns>True if the document is checked out, otherwise false.</returns>
    ''' <remarks></remarks>
    Function IsCheckedOut(ByVal lpId As String) As Boolean

    ''' <summary>
    ''' Deletes a document from the repository
    ''' </summary>
    ''' <param name="lpId">The document identifier</param>
    ''' <returns>True if successful, false if failed</returns>
    ''' <remarks>If the caller does not have sufficient rights to delete the document, an InsufficientRightsException will be thrown</remarks>
    Function DeleteDocument(ByVal lpId As String) As Boolean

    ''' <summary>
    ''' Deletes a version of a document from the repository
    ''' </summary>
    ''' <param name="lpDocumentId">The document identifier</param>
    ''' <param name="lpCriterion">The deletion criterion</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function DeleteVersion(ByVal lpDocumentId As String, ByVal lpCriterion As String) As Boolean

    ''' <summary>
    ''' Updates the properties of the document
    ''' </summary>
    ''' <param name="Args">Contains the document identifier as well as an array of properties to update</param>
    ''' <returns>True if successful, false if failed</returns>
    ''' <remarks>If the caller does not have sufficient rights to update the document, an InsufficientRightsException will be thrown</remarks>
    Function UpdateDocumentProperties(ByVal Args As DocumentPropertyArgs) As Boolean

    ''' <summary>
    ''' Files the document into the specified folder
    ''' </summary>
    ''' <param name="lpId">The document identifier</param>
    ''' <param name="lpFolderPath">The fully qualified repository folder path to file the document into, if the folder path does not exist, the method will attempt to create it.</param>
    ''' <returns>True if successful, false if failed</returns>
    ''' <remarks>If the caller does not have sufficient rights to file the document, an InsufficientRightsException will be thrown</remarks>
    Function FileDocument(ByVal lpId As String, ByVal lpFolderPath As String) As Boolean

    ''' <summary>
    ''' Un-Files a document from the specified folder
    ''' </summary>
    ''' <param name="lpId">The document identifier</param>
    ''' <param name="lpFolderPath">The fully qualified repository folder path to un-file the document from</param>
    ''' <returns>True if successful, false if failed</returns>
    ''' <remarks>If the caller does not have sufficient rights to un-file the document, 
    ''' an InsufficientRightsException will be thrown.  
    ''' If the folder path does not exist or the document is not currently filed in 
    ''' the specified folder path, an InvalidPathException will be thrown.</remarks>
    Function UnFileDocument(ByVal lpId As String, ByVal lpFolderPath As String) As Boolean

#End Region

  End Interface

  ''' <summary>
  ''' Interface to be implemented by all providers that export documents.
  ''' </summary>
  ''' <remarks></remarks>
  Public Interface IDocumentExporter

#Region "Events"

    ''' <summary>
    ''' Fired when a document has been successfully been exported from the source repository.
    ''' </summary>
    ''' <remarks></remarks>
    Event DocumentExported As DocumentExportEventHandler
    Event FolderDocumentExported As FolderDocumentExportEventHandler

    ''' <summary>
    ''' Fired when a folder has been successfully been exported from the source repository.
    ''' </summary>
    ''' <remarks></remarks>
    Event FolderExported As FolderExportedEventHandler

    ''' <summary>
    ''' Fired when a document export has failed.
    ''' </summary>
    ''' <remarks></remarks>
    Event DocumentExportError As DocumentExportErrorEventHandler

    ''' <summary>
    ''' Fired when the provider wishes to display an informational message
    ''' </summary>
    ''' <remarks></remarks>
    Event DocumentExportMessage As DocumentExportMessageEventHandler

#End Region

#Region "Properties"

    ''' <summary>
    ''' Gets or sets the default path to which the provider will export documents.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property ExportPath() As String

#End Region

#Region "Methods"

#Region "Event Handler Methods"

    Sub OnDocumentExported(ByRef e As DocumentExportedEventArgs)
    Sub OnFolderDocumentExported(ByRef e As FolderDocumentExportedEventArgs)
    Sub OnFolderExported(ByRef e As FolderExportedEventArgs)
    Sub OnDocumentExportError(ByRef e As DocumentExportErrorEventArgs)
    Sub OnDocumentExportMessage(ByRef e As WriteMessageArgs)

#End Region

#Region "Methods"

    ''' <summary>
    ''' Gets the count of documents in the specified folder path
    ''' </summary>
    ''' <param name="lpFolderPath"></param>
    ''' <param name="lpRecursionLevel"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function DocumentCount(ByVal lpFolderPath As String, Optional ByVal lpRecursionLevel As Core.RecursionLevel = Core.RecursionLevel.ecmThisLevelOnly) As Long

    ''' <summary>
    ''' Exports the document specified by ID
    ''' </summary>
    ''' <param name="lpId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function ExportDocument(ByVal lpId As String) As Boolean

    ''' <summary>
    ''' Exports a document using the specified argument object
    ''' </summary>
    ''' <param name="Args"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function ExportDocument(ByVal Args As ExportDocumentEventArgs) As Boolean

    ' ''' <summary>
    ' ''' Exports a collection of documents using the specified argument object
    ' ''' </summary>
    ' ''' <param name="Args"></param>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    '<Obsolete("This method is no longer supported as of CTS 3.5.", True)>
    'Function ExportDocuments(ByVal Args As ExportDocumentsEventArgs) As Boolean

    ' <Removed by: Ernie at: 9/29/2014-1:57:37 PM on machine: ERNIE-THINK>
    '     ''' <summary>
    '     ''' Exports a folder using the specified argument object.
    '     ''' </summary>
    '     ''' <param name="Args"></param>
    '     ''' <remarks></remarks>
    '     Sub ExportFolder(ByVal Args As ExportFolderEventArgs)
    ' </Removed by: Ernie at: 9/29/2014-1:57:37 PM on machine: ERNIE-THINK>
    'Function GetDocument(ByVal lpId As String) As Document

    ' <Removed by: Ernie at: 9/29/2014-2:07:21 PM on machine: ERNIE-THINK>
    '     ''' <summary>
    '     ''' Sets a document as read-only in the source repository.
    '     ''' </summary>
    '     ''' <param name="lpId"></param>
    '     ''' <returns></returns>
    '     ''' <remarks></remarks>
    '     Function SetDocumentAsReadOnly(ByVal lpId As String) As Boolean
    ' </Removed by: Ernie at: 9/29/2014-2:07:21 PM on machine: ERNIE-THINK>

#End Region

#End Region

  End Interface

  Public Interface IFolderExporter

#Region "Events"

    ''' <summary>
    ''' Fired when a folder export has failed.
    ''' </summary>
    ''' <remarks></remarks>
    Event FolderExportError As FolderExportErrorEventHandler

#End Region

    ''' <summary>
    ''' Exports a folder instance
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>This is not to be used for exporting documents recursively.  This is to be used to export folder objects themselves.</remarks>
    Function ExportFolder(ByVal Args As ExportFolderEventArgs) As Boolean

  End Interface

  Public Interface IFolderImporter

    ''' <summary>
    ''' Imports a folder instance
    ''' </summary>
    ''' <param name="Args"></param>
    ''' <remarks>This is not to be used for importing documents recursively.  This is to be used to import folder objects themselves.</remarks>
    Function ImportFolder(Args As ImportFolderArgs) As Boolean

  End Interface

  Public Interface IAnnotationExporter

#Region "Events"

    Event AnnotationExported As AnnotationExportEventHandler
    Event AnnotationExportError As AnnotationExportEventHandler
    Event AnnotationExportMessage As AnnotationExportMessageEventHandler

#End Region

#Region "Methods"

#Region "Event Handler Methods"

    Sub OnAnnotationExported(ByRef e As AnnotationExportedEventArgs)
    Sub OnAnnotationExportError(ByRef e As AnnotationExportErrorEventArgs)

#End Region

#Region "Methods"

    Function ExportAnnotation(ByVal lpId As String) As Boolean

    Function ExportAnnotation(ByVal Args As ExportAnnotationEventArgs) As Boolean

    '    Function ExportAnnotations(ByVal Args As ExportAnnotationsEventArgs) As Boolean

#End Region

#End Region

  End Interface

  ''' <summary>
  ''' Interface to be implemented by all providers that import DocumentClasses
  ''' </summary>
  ''' <remarks></remarks>
  Public Interface IDocumentClassImporter

    Event DocumentClassImported As DocumentClassImportedEventHandler
    Function DocumentClassExists(ByRef lpId As ObjectIdentifier, Optional ByRef lpReturnedObjectId As String = "") As Boolean
    Function ImportDocumentClass(ByRef Args As ImportDocumentClassEventArgs) As Boolean

  End Interface

  ''' <summary>
  ''' Interface to be implemented by all providers that import documents.
  ''' </summary>
  ''' <remarks></remarks>
  Public Interface IDocumentImporter

#Region "Events"

    ''' <summary>
    ''' Fired when a document has been successfully imported into the target repository.
    ''' </summary>
    ''' <remarks></remarks>
    Event DocumentImported As DocumentImportedEventHandler

    ''' <summary>
    ''' Fired when a document import has failed.
    ''' </summary>
    ''' <remarks></remarks>
    Event DocumentImportError As DocumentImportErrorEventHandler

    '''' <summary>
    '''' Fired when the provider wishes to display an informational message
    '''' </summary>
    '''' <remarks></remarks>
    Event DocumentImportMessage As DocumentImportMessageEventHandler

#End Region

#Region "Event Handler Methods"

    Sub OnDocumentImported(ByRef e As DocumentImportedEventArgs)
    Sub OnDocumentImportError(ByRef e As DocumentImportErrorEventArgs)

#End Region

    ' <Removed by: Ernie at: 9/29/2014-11:16:05 AM on machine: ERNIE-THINK>
    '     ''' <summary>
    '     ''' Gets or sets the default path from which the provider will import documents.
    '     ''' </summary>
    '     ''' <value></value>
    '     ''' <returns></returns>
    '     ''' <remarks></remarks>
    '     Property ImportPath() As String
    ' </Removed by: Ernie at: 9/29/2014-11:16:05 AM on machine: ERNIE-THINK>

    ''' <summary>
    ''' Gets a value that determines whether or not import 
    ''' operations will only allow import of documents using 
    ''' classes defined for the destination repository
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Only applicable if the import provider also implements IClassification</remarks>
    ReadOnly Property EnforceClassificationCompliance() As Boolean

    ''' <summary>
    ''' Imports a document using the specified argument object.
    ''' </summary>
    ''' <param name="Args"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function ImportDocument(ByRef Args As ImportDocumentArgs) As Boolean

  End Interface

  Public Interface IFolderManager

    ''' <summary>
    ''' Checks to see if a folder exists with the specified path.
    ''' </summary>
    ''' <param name="lpFolderPath">The folder path to test.</param>
    ''' <returns>True if the folder exists, otherwise false.</returns>
    ''' <remarks></remarks>
    Function FolderPathExists(lpFolderPath As String) As Boolean

    ''' <summary>
    ''' Checks to see if a folder exists with the specified id.
    ''' </summary>
    ''' <param name="lpId">The folder id to test.</param>
    ''' <returns>True if the folder exists, otherwise false.</returns>
    ''' <remarks></remarks>
    Function FolderIDExists(lpId As String) As Boolean

    ''' <summary>
    ''' Checks to see if a folder exists with the given parent folder id and folder name.
    ''' </summary>
    ''' <param name="lpParentId">The id of the parent folder.</param>
    ''' <param name="lpName">The name of the folder to test.</param>
    ''' <returns>True if the folder exists, otherwise false.</returns>
    ''' <remarks></remarks>
    Function FolderExists(lpParentId As String, lpName As String, ByRef lpExistingFolderId As String) As Boolean

    ''' <summary>
    ''' Creates a folder using the specified path.
    ''' </summary>
    ''' <param name="lpFolderPath">The folder path to create.</param>
    ''' <returns>Returns an IFolder reference to the newly created folder.</returns>
    ''' <remarks></remarks>
    ''' <exception cref="Exceptions.FolderAlreadyExistsException">
    ''' If a folder already exists with the given path, 
    ''' a FolderAlreadyExistsException will be thrown.  
    ''' To avoid this check for the existence of the folder 
    ''' first using FolderPathExists.
    ''' </exception>
    Function CreateFolder(lpFolderPath As String) As IFolder

    ''' <summary>
    ''' Creates a new folder using the specified parameters.
    ''' </summary>
    ''' <param name="lpParentId">The system identifier of the parent folder.</param>
    ''' <param name="lpName">The name of the new folder.</param>
    ''' <param name="lpClassName">
    ''' The folder class for the new folder, if applicable.  If a null value is specified then the default folder class will be used.
    ''' </param>
    ''' <param name="lpProperties">
    ''' A collection of properties to set for the new folder.
    ''' </param>
    ''' <param name="lpPermissions">
    ''' NOT YET IMPLEMENTED, will be used to specify security ACL's for the new folder.
    ''' </param>
    ''' <returns>Returns an IFolder reference to the newly created folder</returns>
    ''' <remarks></remarks>
    ''' <exception cref="Exceptions.FolderAlreadyExistsException">
    ''' If a folder already exists with the given parameters, 
    ''' a FolderAlreadyExistsException will be thrown.  
    ''' To avoid this check for the existence of the folder 
    ''' first using FolderExists.
    ''' </exception>
    Function CreateFolder(lpParentId As String, lpName As String,
                          lpClassName As String,
                          lpProperties As IProperties,
                          lpPermissions As IPermissions) As IFolder

    ''' <summary>
    ''' Deletes the folder at the specified path.
    ''' </summary>
    ''' <param name="lpFolderPath">The path of the folder to delete.</param>
    ''' <remarks></remarks>
    ''' <exception cref="Exceptions.FolderDoesNotExistException">
    ''' If a folder does not exist with the given path, 
    ''' a FolderDoesNotExistException will be thrown.  
    ''' To avoid this check for the existence of the folder 
    ''' first using FolderPathExists.
    ''' </exception>
    Sub DeleteFolderByPath(lpFolderPath As String)

    ''' <summary>
    ''' Deletes the folder by the specified id.
    ''' </summary>
    ''' <param name="lpId">The id of the folder to delete.</param>
    ''' <remarks></remarks>
    ''' <exception cref="Exceptions.FolderDoesNotExistException">
    ''' If a folder does not exist with the given id, 
    ''' a FolderDoesNotExistException will be thrown.  
    ''' To avoid this check for the existence of the folder 
    ''' first using FolderIDExists.
    ''' </exception>
    Sub DeleteFolderByID(lpId As String)

    ''' <summary>
    ''' Files the document in the folder specified.
    ''' </summary>
    ''' <param name="lpDocumentID">The system document identifier.</param>
    ''' <param name="lpFolderID">The target system folder identifier.</param>
    ''' <remarks></remarks>
    ''' <exception cref="Exceptions.FolderDoesNotExistException">
    ''' If no folder exists by the specified id a FolderDoesNotExistException will be thrown.
    ''' </exception>
    ''' <exception cref="Exceptions.DocumentDoesNotExistException">
    ''' If no document exists by the specified id a DocumentDoesNotExistException will be thrown.
    ''' </exception>
    Sub FileDocumentByID(lpDocumentID As String, lpFolderID As String)

    ''' <summary>
    ''' Files the document in the folder specified.
    ''' </summary>
    ''' <param name="lpDocumentID">The system document identifier.</param>
    ''' <param name="lpFolderPath">The target folder path.</param>
    ''' <remarks></remarks>
    ''' <exception cref="Exceptions.FolderDoesNotExistException">
    ''' If no folder exists by the specified path a FolderDoesNotExistException will be thrown.
    ''' </exception>
    ''' <exception cref="Exceptions.DocumentDoesNotExistException">
    ''' If no document exists by the specified id a DocumentDoesNotExistException will be thrown.
    ''' </exception>
    Sub FileDocumentByPath(lpDocumentID As String, lpFolderPath As String)

    ''' <summary>
    ''' Unfiles the document from the folder specified.
    ''' </summary>
    ''' <param name="lpDocumentID">The system document identifier.</param>
    ''' <param name="lpFolderID">The current system folder identifier.</param>
    ''' <remarks></remarks>
    ''' <exception cref="Exceptions.FolderDoesNotExistException">
    ''' If no folder exists by the specified id a FolderDoesNotExistException will be thrown.
    ''' </exception>
    ''' <exception cref="Exceptions.DocumentDoesNotExistException">
    ''' If no document exists by the specified id a DocumentDoesNotExistException will be thrown.
    ''' </exception>
    Sub UnfileDocumentByID(lpDocumentID As String, lpFolderID As String)

    ''' <summary>
    ''' Unfiles the document from the folder specified.
    ''' </summary>
    ''' <param name="lpDocumentID">The system document identifier.</param>
    ''' <param name="lpFolderPath">The current folder path.</param>
    ''' <remarks></remarks>
    ''' <exception cref="Exceptions.FolderDoesNotExistException">
    ''' If no folder exists by the specified path a FolderDoesNotExistException will be thrown.
    ''' </exception>
    ''' <exception cref="Exceptions.DocumentDoesNotExistException">
    ''' If no document exists by the specified id a DocumentDoesNotExistException will be thrown.
    ''' </exception>
    Sub UnfileDocumentByPath(lpDocumentID As String, lpFolderPath As String)

    ''' <summary>
    ''' Moves the document from the current specified folder to the new folder.
    ''' </summary>
    ''' <param name="lpDocumentID">The system document identifier.</param>
    ''' <param name="lpCurrentFolderID">The current system folder identifier.</param>
    ''' <param name="lpDestinationFolderID">The target system folder identifier.</param>
    ''' <remarks></remarks>
    ''' <exception cref="Exceptions.FolderDoesNotExistException">
    ''' If no folder exists by the specified current or new folder id a 
    ''' FolderDoesNotExistException will be thrown.
    ''' </exception>
    ''' <exception cref="Exceptions.DocumentDoesNotExistException">
    ''' If no document exists by the specified id a DocumentDoesNotExistException will be thrown.
    ''' </exception>
    Sub MoveDocumentByID(lpDocumentID As String, lpCurrentFolderID As String, lpDestinationFolderID As String)

    ''' <summary>
    ''' Moves the document from the current specified folder to the new folder.
    ''' </summary>
    ''' <param name="lpDocumentID">The system document identifier.</param>
    ''' <param name="lpCurrentFolderPath">The current system folder path.</param>
    ''' <param name="lpDestinationFolderID">The target system folder path.</param>
    ''' <remarks></remarks>
    ''' <exception cref="Exceptions.FolderDoesNotExistException">
    ''' If no folder exists by the specified current or new folder path a 
    ''' FolderDoesNotExistException will be thrown.
    ''' </exception>
    ''' <exception cref="Exceptions.DocumentDoesNotExistException">
    ''' If no document exists by the specified id a DocumentDoesNotExistException will be thrown.
    ''' </exception>
    Sub MoveDocumentByPath(lpDocumentID As String, lpCurrentFolderPath As String, lpDestinationFolderID As String)

    ''' <summary>
    ''' Gets the Id of a folder for the specified path.
    ''' </summary>
    ''' <param name="lpFolderPath">The path of the folder to get.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <exception cref="Exceptions.FolderDoesNotExistException">
    ''' If a folder does not exist with the given path, 
    ''' a FolderDoesNotExistException will be thrown.  
    ''' To avoid this check for the existence of the folder 
    ''' first using FolderPathExists.
    ''' </exception>
    Function GetFolderIdByPath(lpFolderPath As String) As String

    ''' <summary>
    ''' Gets a reference to the folder using the specified path.
    ''' </summary>
    ''' <param name="lpFolderPath">The path of the folder to get.</param>
    ''' <param name="lpMaxContentCount">The maximum number of content elements to return.  -1 will return all possible elements.</param>
    ''' <returns>An IFolder object reference to the folder using the specified path.</returns>
    ''' <remarks></remarks>
    ''' <exception cref="Exceptions.FolderDoesNotExistException">
    ''' If a folder does not exist with the given path, 
    ''' a FolderDoesNotExistException will be thrown.  
    ''' To avoid this check for the existence of the folder 
    ''' first using FolderPathExists.
    ''' </exception>
    Function GetFolderInfoByPath(lpFolderPath As String, lpMaxContentCount As Integer) As IFolder

    ''' <summary>
    ''' Gets a reference to the folder using the specified id.
    ''' </summary>
    ''' <param name="lpId">The id of the folder to get.</param>
    ''' <param name="lpMaxContentCount">The maximum number of content elements to return.  -1 will return all possible elements.</param>
    ''' <returns>An IFolder object reference to the folder using the specified path.</returns>
    ''' <remarks></remarks>
    ''' <exception cref="Exceptions.FolderDoesNotExistException">
    ''' If a folder does not exist with the given id, 
    ''' a FolderDoesNotExistException will be thrown.  
    ''' To avoid this check for the existence of the folder 
    ''' first using FolderIDExists.
    ''' </exception>
    Function GetFolderInfoByID(lpId As String, lpMaxContentCount As Integer) As IFolder

    Function UpdateFolderProperties(ByVal Args As FolderPropertyArgs) As Boolean

  End Interface

  Public Interface ILinkManager

    ''' <summary>
    ''' Create a document link and returns the id of the newly created link.
    ''' </summary>
    ''' <param name="Args"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CreateDocumentLink(ByVal Args As CreateDocumentLinkEventArgs) As String

    ' ''' <summary>
    ' ''' Replaces an existing document with a document link and returns the id of the newly created link.
    ' ''' </summary>
    ' ''' <param name="lpOriginalId"></param>
    ' ''' <param name="Args"></param>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Function ReplaceItemWithLink(ByVal lpOriginalId As String, ByVal Args As CreateDocumentLinkEventArgs) As String

  End Interface

  Public Interface IAnnotationImporter

#Region "Events"

    'Event AnnotationImported As AnnotationImportEventHandler
    'Event AnnotationImportError As AnnotationImportEventHandler
    'Event AnnotationImportMessage As AnnotationImportMessageEventHandler

#End Region

#Region "Event Handler Methods"

    'Sub OnAnnotationImported(ByRef e As AnnotationImportedEventArgs)
    'Sub OnAnnotationImportError(ByRef e As AnnotationImportErrorEventArgs)

#End Region

    'Function ImportAnnotation(ByRef Args As ImportAnnotationEventArgs) As Boolean

  End Interface

  Public Interface IRepositoryDiscovery

#Region "Methods"

    Function GetRepositories() As RepositoryIdentifiers

#End Region

  End Interface

  ''' <summary>Interface to be implemented by all providers.</summary>
  ''' <remarks>Implemented by the base CProvider class.</remarks>
  Public Interface IProvider
    Inherits IDisposable

#Region "Public Events"

    Event ConnectionStateChanged As ConnectionStateChangedEventHandler

#Region "Background Worker Events"

    Event BackgroundWorker_Disposed As EventHandler
    Event BackgroundWorker_DoWork As System.ComponentModel.DoWorkEventHandler
    Event BackgroundWorker_ProgressChanged As System.ComponentModel.ProgressChangedEventHandler
    Event BackgroundWorker_RunWorkerCompleted As System.ComponentModel.RunWorkerCompletedEventHandler

#End Region

#Region "ProviderProperty Events"

    Event ProviderProperty_ValueChanged As ProviderPropertyValueChangedEventHandler

#End Region

#End Region

#Region "Properties"

    Property ConnectionString() As String
    Property ContentSource() As ContentSource
    Property SelectedFolder() As IFolder
    Property UserName As String
    Property Password As String
    Property Tag As Object
    ReadOnly Property HasValidLicense() As Boolean
    ReadOnly Property LicenseFailureReason() As String
    ReadOnly Property State As ProviderConnectionState
    'ReadOnly Property Feature As FeatureEnum
    ReadOnly Property TokenRequired As Boolean

#Region "Read Only Properties"

    ReadOnly Property Name() As String
    ReadOnly Property ProviderClass() As ProviderClass
    ReadOnly Property ProviderProperties() As ProviderProperties
    ReadOnly Property System() As ProviderSystem
    ReadOnly Property Search() As ISearch
    ReadOnly Property Information() As ProviderInformation
    ReadOnly Property ActionProperties() As ActionProperties
    ReadOnly Property IsConnected() As Boolean
    ReadOnly Property IsInitialized() As Boolean
    ReadOnly Property FolderDelimiter() As String
    ReadOnly Property LeadingFolderDelimiter() As Boolean

#End Region

#End Region

#Region "Methods"

    Sub Connect()
    Sub Connect(ByVal ConnectionString As String)
    Sub Connect(ByVal ContentSource As ContentSource)
    Sub Connect(ByVal Connection As IRepositoryConnection)
    Sub Disconnect()
    Function GenerateConnectionString() As String
    Function GenerateConnectionString(ByVal lpProviderProperties As ProviderProperties) As String
    Function GenerateConnectionString(ByVal lpContentSourceName As String, ByVal lpProviderProperties As ProviderProperties) As String
    Function GenerateConnectionString(ByVal lpIncludeProviderPath As Boolean) As String
    Function GenerateConnectionString(ByVal lpProviderProperties As ProviderProperties,
                                      ByVal lpIncludeProviderPath As Boolean) As String
    Function GenerateConnectionString(ByVal lpContentSourceName As String,
                                      ByVal lpProviderProperties As ProviderProperties,
                                      ByVal lpIncludeProviderPath As Boolean) As String

    ''' <summary>
    ''' Gets the folder based on the path.
    ''' </summary>
    ''' <param name="lpFolderPath">The fully qualifed path of the folder</param>
    ''' <param name="lpMaxContentCount">Pass -1 to get all content, 0 to get no content, 
    ''' or any other number to specify a specific maximum</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function GetFolder(ByVal lpFolderPath As String, ByVal lpMaxContentCount As Long) As IFolder
    Function SupportsInterface(ByVal lpInterface As ProviderClass) As Boolean
    Function GetDocumentByPath(lpDocumentPath As String) As Document
    Function GetInterface(ByVal lpInterface As ProviderClass) As Object

    ' <Added by: Ernie at: 9/28/2011-8:12:28 AM on machine: ERNIE-M4400>
    ''' <summary>
    ''' Forces the initialization of the classification related properties so 
    ''' that they may be initialized in advance instead of upon first request.
    ''' </summary>
    ''' <remarks></remarks>
    Sub InitializeClassificationProperties()
    ' </Added by: Ernie at: 9/28/2011-8:12:28 AM on machine: ERNIE-M4400>

    Sub InitializeProvider(ByVal lpContentSource As ContentSource)
    Function CreateSearch() As ISearch
    ''' <summary>
    ''' Gets a list of available values for the specified provider property.
    ''' </summary>
    ''' <param name="lpProviderProperty">The provider property to
    '''  which the values will be made available.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function GetAvailableValues(ByVal lpProviderProperty As ProviderProperty) As IEnumerable(Of String)

    Function ReadyToGetAvailableValues(ByVal lpProviderProperty As ProviderProperty) As Boolean

#End Region

  End Interface

  'Public Interface IProviderImages
  '  ReadOnly Property Images() As ImageSet
  'End Interface

  Public Interface ICustomObject
    Inherits IRepositoryObject
    Inherits IMetaHolder
    Inherits IAuditableItem

#Region "Properties"


    ''' <summary>
    ''' The name of the class of object.
    ''' </summary>
    ''' <returns></returns>
    Property ClassName As String

    ''' <summary>
    ''' Reference to the parent Content Source to which this object belongs
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property ContentSource() As ContentSource

    ''' <summary>
    ''' Reference to the parent Provider to which this object belongs
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property Provider() As IProvider

#End Region

  End Interface

  Public Interface ICustomObjectExporter

#Region "Events"

    ''' <summary>
    ''' Fired when a object export has failed.
    ''' </summary>
    ''' <remarks></remarks>
    Event ObjectExportError As ObjectExportErrorEventHandler

#End Region

    ''' <summary>
    ''' Exports a object instance
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>This is to be used to export custom objects.</remarks>
    Function ExportObject(ByVal Args As ExportObjectEventArgs) As Boolean

  End Interface

  Public Interface ICustomObjectImporter

    ''' <summary>
    ''' Imports a custom object instance
    ''' </summary>
    ''' <param name="Args"></param>
    ''' <remarks>This is to be used to import custom objects.</remarks>
    Function ImportObject(Args As ImportObjectArgs) As Boolean

  End Interface

  Public Interface IFolder
    Inherits IRepositoryObject
    Inherits IMetaHolder
    Inherits IAuditableItem

#Region "Properties"

    ''' <summary>
    ''' Fired when a folder has been selected.
    ''' </summary>
    ''' <remarks></remarks>
    Event FolderSelected As FolderEventHandler

    Sub OnFolderSelected()

    ''' <summary>
    ''' The label for the folder
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Typically used for the folder label by the tree view</remarks>
    Property Label() As String

    ''' <summary>
    ''' Reference to the parent Content Source to which this folder belongs
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property ContentSource() As ContentSource

    ''' <summary>
    ''' Reference to the parent Provider to which this folder belongs
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property Provider() As IProvider

    ''' <summary>
    ''' Specifies whether or not there is any actual root folder in the underlying provider
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property InvisiblePassThrough() As Boolean

    ''' <summary>
    ''' Gets or Sets the maximum number of contents to be retrieved for the folder
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property MaxContentCount() As Long

    Overloads Property Name As String

#Region "Read Only Properties"

    ReadOnly Property ContentCount() As Long
    ReadOnly Property ContentCount(ByVal FolderRecursionLevel As Core.RecursionLevel) As Long
    ReadOnly Property Contents() As FolderContents
    ReadOnly Property FolderNames() As Collections.Specialized.StringCollection
    ReadOnly Property HasContent(ByVal lpRecursionLevel As Core.RecursionLevel) As Boolean
    ReadOnly Property HasSubFolders() As Boolean
    ReadOnly Property Path() As String
    ReadOnly Property SubFolderCount() As Long
    ReadOnly Property SubFolders() As Folders
    Function GetSubFolders(ByVal lpGetContents As Boolean) As Folders
    ReadOnly Property TreeLabel() As String

#End Region

#End Region

#Region "Methods"

    Function GetID() As String
    Function GetPath() As String
    Function GetSubFolderCount() As Long
    Function GetFolderByPath(ByVal lpFolderPath As String, ByVal lpMaxContentCount As Long) As IFolder
    Function GetContentCount() As Long
    Function GetContentCount(ByVal lpFolderRecursionLevel As Core.RecursionLevel) As Long
    Sub InitializeFolderCollection(ByVal lpFolderPath As String)
    ''' <summary>
    ''' Gets the content for the current folder and refreshes the subfolder collection.
    ''' </summary>
    ''' <remarks></remarks>
    Sub Refresh()

#End Region

  End Interface

  Public Interface ISearch

#Region "Events"
    Event SearchUpdate As SearchUpdateEventHandler
    Event SearchComplete As SearchCompleteEventHandler
#End Region

#Region "Properties"

    Property Criteria() As Criteria
    Property Provider() As IProvider

#Region "Read Only Properties"

    ReadOnly Property DefaultResultColumns As ResultColumns
    ReadOnly Property DefaultQueryTarget As String
    ReadOnly Property SearchResultSet() As SearchResultSet
    ReadOnly Property DataSource() As Data.DataSource

#End Region

#End Region

#Region "Methods"

    Function Execute(ByVal Args As SearchArgs) As SearchResultSet
    Function Execute() As SearchResultSet
    'Function SimpleSearch(ByVal args As SimpleSearchArgs) As SearchResultSet
    Function SimpleSearch(ByVal args As SimpleSearchArgs) As DataTable
    ''' <summary>
    ''' Resets the current search
    ''' </summary>
    ''' <remarks></remarks>
    Sub Reset()

#End Region

  End Interface

  Public Interface ISQLPassThroughSearch
    Function Execute(ByVal Sql As String) As SearchResultSet
  End Interface

  ''' <summary>
  ''' Interface to be implemented by all providers that integrate into the Ecmg
  ''' Explorer application.
  ''' </summary>
  Public Interface IExplorer
    ReadOnly Property RootFolder() As IFolder
    'ReadOnly Property SelectedFolder() As IFolder
    ReadOnly Property Search() As ISearch
    ReadOnly Property GetFolderByID(ByVal lpFolderID As String, ByVal lpFolderLevels As Int32, ByVal lpMaxContentCount As Int32) As IFolder
    ReadOnly Property GetFolderContentsByID(ByVal lpFolderID As String, ByVal lpMaxContentCount As Int32) As FolderContents
    ReadOnly Property HasSubFolders(ByVal lpFolderID As String) As Boolean
    ReadOnly Property IsFolderValid(ByVal lpFolderID As String) As Boolean
  End Interface

  ''' <summary>
  ''' Interface to be implemented by all providers that provide repository
  ''' classification information.
  ''' </summary>
  Public Interface IClassification
    ReadOnly Property ContentProperties() As Core.ClassificationProperties
    ReadOnly Property DocumentClasses() As Core.DocumentClasses
    ReadOnly Property DocumentClass(ByVal lpDocumentClassName As String) As Core.DocumentClass

  End Interface

  ''' <summary>
  ''' Interface to be implemented by all providers that provide repository
  ''' custom object classification information.
  ''' </summary>
  Public Interface ICustomObjectClassification
    ReadOnly Property ObjectProperties() As Core.ClassificationProperties
    ReadOnly Property ObjectClasses() As Core.ObjectClasses
    ReadOnly Property ObjectClass(ByVal lpObjectClassName As String) As Core.ObjectClass

  End Interface

  ''' <summary>
  ''' Interface to be implemented by all providers that provide repository
  ''' folder classification information.
  ''' </summary>
  Public Interface IFolderClassification
    ReadOnly Property FolderProperties() As Core.ClassificationProperties
    ReadOnly Property FolderClasses() As Core.FolderClasses
    ReadOnly Property FolderClass(ByVal lpFolderClassName As String) As Core.FolderClass

  End Interface

  ''' <summary>
  ''' Interface to be implemented by all providers that provide repository 
  ''' security classification information.
  ''' </summary>
  ''' <remarks></remarks>
  Public Interface ISecurityClassification
    ReadOnly Property AvailableRights As Security.IAccessRights
  End Interface

  Public Interface IUpdatePermissions
    Function UpdatePermissions(ByVal Args As Arguments.ObjectSecurityArgs) As Boolean
  End Interface

  ''' <summary>
  ''' Interface to be implemented by all providers that provides a way to retrieve a document. 
  ''' </summary>
  ''' <remarks></remarks>
  Public Interface ICopy
    Function GetDocument(ByVal lpId As String, ByVal lpGetContents As Boolean, ByVal lpStorageType As Content.StorageTypeEnum, ByVal lpDestinationFolder As String, ByVal SessionId As String) As Document
    Function GetDocumentXml(ByVal lpId As String, ByVal lpGetContents As Boolean, ByVal lpStorageType As Content.StorageTypeEnum, ByVal lpDestinationFolder As String, ByVal SessionId As String) As String
  End Interface

  ''' <summary>
  ''' Interface to be implemented by all providers that provides a way to create an instance of
  ''' a document and add a new document
  ''' </summary>
  ''' <remarks></remarks>
  Public Interface ICreate
    Function CreateDocumentInstance(ByVal lpDocumentClassName As String, ByVal lpSessionId As String) As Document
    Function AddDocument(ByVal lpDocument As Core.Document, ByVal lpFolderPath As String, ByVal lpSessionId As String) As String '(returns the new ID)
  End Interface

  ''' <summary>
  ''' Interface to be implemented by all providers that provides a way to delete a document
  ''' </summary>
  ''' <remarks></remarks>
  Public Interface IDelete
    Function DeleteDocument(ByVal lpId As String, ByVal lpSessionId As String) As Boolean
    ' <Added by: Ernie at: 1/12/2013-7:17:46 AM on machine: ERNIE-THINK>
    Function DeleteDocument(ByVal lpId As String, ByVal lpSessionId As String, ByVal lpDeleteAllVersions As Boolean) As Boolean
    ' </Added by: Ernie at: 1/12/2013-7:17:46 AM on machine: ERNIE-THINK>
    Function DeleteVersion(ByVal lpDocumentId As String, ByVal lpCriterion As String, ByVal lpSessionId As String) As Boolean
  End Interface

  ''' <summary>
  ''' Interface to be implemented by all providers that provides a way to update the properties a document
  ''' </summary>
  ''' <remarks></remarks>
  Public Interface IUpdateProperties
    'UpdateDocumentProperties(DocumentId, ECMProperty(), SessionId) as bool
    Function UpdateDocumentProperties(ByVal Args As Arguments.DocumentPropertyArgs) As Boolean
  End Interface

  ''' <summary>
  ''' Interface to be implemented by all providers that provides a way to file/unfile a document
  ''' </summary>
  ''' <remarks></remarks>
  Public Interface IFile
    Function FileDocument(ByVal lpId As String, ByVal lpFolderPath As String, ByVal lpSessionId As String) As Boolean
    Function UnFileDocument(ByVal lpId As String, ByVal lpFolderPath As String, ByVal lpSessionId As String) As Boolean
  End Interface

  ''' <summary>
  ''' Interface to be implemented by all providers that provides a way to rename the title of a document
  ''' </summary>
  ''' <remarks></remarks>
  Public Interface IRename
    Function RenameDocument(ByVal lpId As String, ByVal lpNewName As String, ByVal SessionId As String) As Boolean
  End Interface

  ''' <summary>
  ''' Interface to be implemented by all providers that provides a way to checkin/checkout/cancel checkout
  ''' for a document
  ''' </summary>
  ''' <remarks></remarks>
  Public Interface IVersion
    Function CancelCheckoutDocument(ByVal lpId As String, ByVal SessionId As String) As Boolean
    Function CheckInDocument(ByVal lpId As String, ByVal lpContentPath As String, ByVal lpAsMajorVersion As Boolean, ByVal lpSessionId As String) As Boolean
    Function CheckInDocument(ByVal lpId As String, ByVal lpContentPath() As String, ByVal lpAsMajorVersion As Boolean, ByVal lpSessionId As String) As Boolean
    Function CheckOutDocument(ByVal lpId As String, ByVal lpDestinationFolder As String, ByRef lpOutputFileNames() As String, ByVal lpSessionID As String) As Boolean
    Function IsCheckedOut(ByVal lpId As String) As Boolean
  End Interface

#End Region

#Region "Provider Enumerations"

  Public Enum ProviderClass
    Classification = -1
    Explorer = 0
    DocumentExporter = 1
    DocumentImporter = 2
    BasicContentServices = 3
    RecordsManager = 4
    ProviderImages = 5
    Copy = 6
    Delete = 7
    Rename = 8
    Version = 9
    File = 10
    UpdateProperties = 11
    Create = 12
    ChoiceListExporter = 13
    ChoiceListImporter = 14
    FolderManager = 15
    FolderClassification = 16
    SecurityClassification = 17
    UpdatePermissions = 18
    SQLPassThroughSearch = 19
    LinkManager = 20
    RepositoryDiscovery = 21
    AnnotationExporter = 22
    AnnotationImporter = 23
    DocumentClassImporter = 26
    FolderExporter = 27
    FolderImporter = 28
    CustomObjectExporter = 29
    CustomObjectImporter = 30
    CustomObjectClassification = 31
  End Enum

  Public Enum MoveType
    Move = 0
    Copy = 1
  End Enum

  Public Enum ProviderConnectionState
    Unavailable = -1
    Disconnected = 0
    Connected = 1
  End Enum

#End Region

#End Region

End Namespace

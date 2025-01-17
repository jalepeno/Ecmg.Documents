'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------


Imports Documents.Arguments

Namespace Core.ChoiceLists

  ''' <summary>Delegate event handler for the ChoiceListImported event.</summary>
  Public Delegate Sub ChoiceListImportedEventHandler(ByVal sender As Object, ByVal e As ChoiceListImportedEventArgs)

  ''' <summary>Delegate event handler for the ChoiceListExported event.</summary>
  Public Delegate Sub ChoiceListExportedEventHandler(ByVal sender As Object, ByVal e As ChoiceListExportedEventArgs)

  ''' <summary>Delegate event handler for the ChoiceListExportError event.</summary>
  Public Delegate Sub ChoiceListExportErrorEventHandler(ByVal sender As Object, ByVal e As ChoiceListExportErrorEventArgs)

  ''' <summary>Delegate event handler for the ChoiceListImportError event.</summary>
  Public Delegate Sub ChoiceListImportErrorEventHandler(ByVal sender As Object, ByVal e As ChoiceListImportErrorEventArgs)

  ''' <summary>
  ''' Interface to be implemented by all providers that import ChoiceLists
  ''' </summary>
  ''' <remarks></remarks>
  Public Interface IChoiceListImporter

#Region "Events"

    Event ChoiceListImported As ChoiceListImportedEventHandler
    ''' <summary>
    ''' Fired when a choicelist import has failed.
    ''' </summary>
    ''' <remarks></remarks>
    Event ChoiceListImportError As ChoiceListImportErrorEventHandler

#End Region

#Region "Event Handler Methods"

    Sub OnChoiceListImported(ByRef e As ChoiceListImportedEventArgs)
    Sub OnChoiceListImportError(ByRef e As ChoiceListImportErrorEventArgs)

#End Region

#Region "Methods"

    Function ChoiceListExists(ByRef lpId As ObjectIdentifier, Optional ByRef lpReturnedObjectId As String = "") As Boolean
    Function ImportChoiceList(ByRef Args As ImportChoiceListEventArgs) As Boolean

#End Region

  End Interface

  Public Interface IChoiceListExporter

#Region "Events"

    Event ChoiceListExported As ChoiceListExportedEventHandler

    ''' <summary>
    ''' Fired when a choicelist export has failed.
    ''' </summary>
    ''' <remarks></remarks>
    Event ChoiceListExportError As ChoiceListExportErrorEventHandler

#End Region

#Region "Event Handler Methods"

    Sub OnChoiceListExported(ByRef e As ChoiceListExportedEventArgs)
    Sub OnChoiceListExportError(ByRef e As ChoiceListExportErrorEventArgs)

#End Region

#Region "Properties"

    ReadOnly Property ChoiceListNames() As List(Of String)

#End Region

#Region "Methods"

    '''' <summary>
    '''' Returns a choice list object using the supplied choice list name
    '''' </summary>
    '''' <param name="lpChoiceListName"></param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    '''' <exception cref="Exceptions.ChoiceListDoesNotExistException">
    '''' If no choice list exists with the given name then 
    '''' a ChoiceListDoesNotExist Exception is thrown
    '''' </exception>
    'Function ExportChoiceList(ByVal lpChoiceListName As String) As ChoiceList

    Function ExportChoiceList(ByRef lpArgs As ExportChoiceListEventArgs) As Boolean

#End Region

  End Interface

#Region "Choice List Object Interfaces"

  Public Interface IChoiceValue
    Inherits IChoiceItem

    Property ChoiceType() As ChoiceType
    Property Value() As Object

  End Interface

  Public Interface IChoiceItem

  End Interface

  Public Interface IChoiceGroup
    Inherits IChoiceItem

    Property ChoiceValues() As IChoiceValues

  End Interface

  Public Interface IChoiceValues
    Inherits ICollection

    Sub Add(ByVal item As IChoiceItem)

    Default Property Item(ByVal index As Integer) As IChoiceItem
    Default ReadOnly Property Item(ByVal name As String) As IChoiceItem

  End Interface

  Public Interface IChoiceList

    Property ChoiceValues() As IChoiceValues

  End Interface

#End Region

End Namespace
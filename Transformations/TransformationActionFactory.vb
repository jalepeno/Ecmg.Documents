' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  TransformationActionFactory.vb
'  Description :  Used for creating operation objects
'  Created     :  11/2/2013 4:16:03 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Globalization
Imports System.IO
Imports System.Reflection
Imports System.Xml
Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Utilities

#End Region

Namespace Transformations

  Public Class TransformationActionFactory

#Region "Class Variables"

    Private Shared mintReferenceCount As Integer
    Private Shared mobjInstance As TransformationActionFactory
    Private mobjAvailableActions As IDictionary(Of String, Type) = Nothing
    Private mobjAvailableActionDisplayNames As IList(Of String) = Nothing
    Private mobjAvailableCoreActions As IDictionary(Of String, Type) = Nothing
    Private mobjAvailableExtensionActions As IDictionary(Of String, Type) = Nothing
    'Private WithEvents mobjExtensionCatalog As ExtensionCatalog = ExtensionCatalog.Instance

#End Region

#Region "Constructors"

    Private Sub New()
      mintReferenceCount = 0
    End Sub

#End Region

#Region "Singleton Support"

    Public Shared Function Instance() As TransformationActionFactory

      Try

        If mobjInstance Is Nothing Then
          mobjInstance = New TransformationActionFactory
        End If

        mintReferenceCount += 1
        Return mobjInstance

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

#End Region

#Region "Public Properties"

    Public ReadOnly Property AvailableActions() As IDictionary(Of String, Type)
      Get

        Try

          If mobjAvailableActions Is Nothing Then
            mobjAvailableActions = GetAllAvailableActions()
          End If

          Return mobjAvailableActions

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try

      End Get
    End Property

    Public ReadOnly Property AvailableActionDisplayNames As IList(Of String)
      Get
        Try
          If mobjAvailableActionDisplayNames Is Nothing Then
            mobjAvailableActionDisplayNames = GetAllAvailableActionDisplayNames()
          End If
          Return mobjAvailableActionDisplayNames
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property AvailableCoreActions As IDictionary(Of String, Type)
      Get
        Try
          If mobjAvailableCoreActions Is Nothing Then
            mobjAvailableCoreActions = GetAvailableCoreActions()
          End If

          Return mobjAvailableCoreActions
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    'Public ReadOnly Property AvailableExtensionActions As IDictionary(Of String, Type)
    '  Get
    '    Try
    '      If mobjAvailableExtensionActions Is Nothing Then
    '        mobjAvailableExtensionActions = GetAvailableExtensionActions()
    '      End If

    '      Return mobjAvailableExtensionActions
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Get
    'End Property

#End Region

#Region "Public Shared Methods"

    Public Shared Function Create(ByVal lpActionType As String) As IActionable

      Try
        Return Instance.CreateAction(lpActionType)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Friend Shared Function Create(lpActionNode As XmlNode,
                              lpLocale As String) As IActionable

      Try
        Return Instance.CreateAction(lpActionNode, lpLocale)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Friend Shared Function Create(ByVal lpAction As IActionable) As IActionable

      Try
        Return Instance.CreateAction(lpAction)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Public Shared Function GetRunOperable(lpRunActionNode As XmlNode) As IActionable

      Try

        If lpRunActionNode.HasChildNodes = False Then
          Return Nothing

        Else

          Return TransformationActionFactory.Create(lpRunActionNode.FirstChild, CultureInfo.CurrentCulture.Name)

        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

#End Region

#Region "Private Methods"

    Private Function GetAllAvailableActions() As IDictionary(Of String, Type)
      Try

        Dim lobjActionList As New SortedDictionary(Of String, Type)
        Dim lobjCoreActionList As IDictionary(Of String, Type) = AvailableCoreActions
        ' Dim lobjExtensionOperationList As IDictionary(Of String, Type) = AvailableExtensionOperations

        For Each lstrCoreOperation As String In lobjCoreActionList.Keys

          If lobjActionList.ContainsKey(lstrCoreOperation) = False Then
            lobjActionList.Add(lstrCoreOperation, lobjCoreActionList.Item(lstrCoreOperation))
          End If

        Next

        'For Each lstrExtensionOperation As String In lobjExtensionOperationList.Keys

        '  If lobjAllOperationList.ContainsKey(lstrExtensionOperation) = False Then
        '    lobjAllOperationList.Add(lstrExtensionOperation, lobjExtensionOperationList.Item(lstrExtensionOperation))
        '  End If

        'Next

        Return lobjActionList

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function GetAllAvailableActionDisplayNames() As IList(Of String)
      Try
        Dim lobjReturnList As New List(Of String)
        For Each lstrKey As String In AvailableActions.Keys()
          lobjReturnList.Add(Helper.CreateDisplayName(lstrKey))
        Next

        Return lobjReturnList

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function GetAvailableCoreActions() As IDictionary(Of String, Type)

      Dim lobjActionList As New SortedDictionary(Of String, Type)
      Dim lobjActionsTypes As List(Of Type)
      Dim lobjAction As IActionableInformation = Nothing

      Try

        lobjActionsTypes = GetAllCoreActionTypes()

        For Each lobjType As Type In lobjActionsTypes

          lobjAction = CType(Activator.CreateInstance(lobjType), IActionableInformation)

          If lobjAction IsNot Nothing Then
            If Not lobjActionList.ContainsKey(lobjAction.Name) Then
              lobjActionList.Add(lobjAction.Name, lobjType)
            End If
          End If

          lobjAction = Nothing

        Next

        Return lobjActionList

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw

      Finally
        ' Clean up references
        lobjActionsTypes = Nothing
      End Try

    End Function

    Public Shared Function GetAvailableActionTypes() As IList(Of String)

      Try

        Dim lobjAvailableActionTypes As New List(Of String)

        For Each lstrKey As String In Instance.GetAvailableCoreActions.Keys
          lobjAvailableActionTypes.Add(lstrKey)
        Next

        lobjAvailableActionTypes.Sort()

        Return lobjAvailableActionTypes

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Private Function GetAllCoreActionTypes() As List(Of Type)

      Dim lobjActionList As New List(Of String)
      Dim lobjActionType As Type = Nothing
      Dim lobjAssembly As Reflection.Assembly = Nothing
      Dim lobjTypes As IEnumerable(Of Type) = Nothing
      Dim lobjCoreOperationTypes As New List(Of Type)
      Dim lobjAction As ITransformationAction = Nothing

      Try

        lobjActionType = GetType(ITransformationAction)
        lobjAssembly = Reflection.Assembly.GetAssembly(lobjActionType)
        lobjTypes = lobjAssembly.GetTypes.Where(Function(t) lobjActionType.IsAssignableFrom(t))
        lobjAction = Nothing

        For Each lobjType As Type In lobjTypes

          If lobjType.IsAbstract Then
            Continue For
          End If

          If lobjType.IsInterface Then
            Continue For
          End If

          lobjAction = CType(Activator.CreateInstance(lobjType), ITransformationAction)

          If lobjAction IsNot Nothing Then
            lobjCoreOperationTypes.Add(lobjType)
          End If

          lobjAction = Nothing

        Next

        Return lobjCoreOperationTypes

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw

      Finally
        ' Clean up references
        lobjAssembly = Nothing
        lobjTypes = Nothing
        lobjActionType = Nothing
      End Try

    End Function

    ''' <summary>Finds all of the valid action extension dlls in the specified folder path.</summary>
    ''' <param name="lpFolderPath">The folder path in which to search for action extensions.</param>
    ''' <returns>A list of valid action extensions.</returns>
    ''' <exception caption="FolderDoesNotExistException" cref="Exceptions.FolderDoesNotExistException">
    ''' If the specified folder does not exist, a FolderDoesNotExistException will be thrown.</exception>
    Public Shared Function GetAvailableExtensions(lpFolderPath As String) As IList(Of String)

      Try

        If Directory.Exists(lpFolderPath) = False Then
          Throw New Exceptions.FolderDoesNotExistException(lpFolderPath,
            String.Format("Unable to get available extensions, the path '{0}' is not valid.", lpFolderPath))
        End If

        Dim lobjExtensionList As New List(Of String)
        Dim lobjAssembly As System.Reflection.Assembly
        Dim lobjExtensionCandidate As Type

        ' Loop through each dll in the folder and see if it is an operation extension.
        For Each lstrExtensionCandidate As String In Directory.GetFiles(lpFolderPath, "*.dll")
          lobjExtensionCandidate = Nothing
          Try
            lobjAssembly = System.Reflection.Assembly.LoadFrom(lstrExtensionCandidate)
            For Each lobjType As Type In lobjAssembly.GetExportedTypes
              If lobjType.IsAbstract = False Then
                lobjExtensionCandidate = lobjType.GetInterface("IActionExtension")
                If lobjExtensionCandidate IsNot Nothing Then
                  If lobjExtensionList.Contains(lstrExtensionCandidate) = False Then
                    lobjExtensionList.Add(lstrExtensionCandidate)
                  End If
                  Continue For
                End If
              End If
            Next
          Catch BadImageEx As BadImageFormatException
            ApplicationLogging.LogException(BadImageEx, Reflection.MethodBase.GetCurrentMethod)
            Continue For
          Catch ReflexLoadEx As ReflectionTypeLoadException
            ApplicationLogging.LogException(ReflexLoadEx, Reflection.MethodBase.GetCurrentMethod)
            Continue For
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            Continue For
          End Try
        Next

        Return lobjExtensionList

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    'Public Shared Function GetAvailableExtensionActions(lpFolderPath As String) As IList(Of IExtensionInformation)

    '  Dim lobjExtensionList As IList(Of String)
    '  Dim lobjExtensionActionList As New List(Of IExtensionInformation)
    '  'Dim lobjExtensionOperations As IEnumerable(Of Type)
    '  Dim lobjExportedTypes As Type() = Nothing
    '  Dim lobjExportedInterfaces As Type() = Nothing
    '  Dim lobjActionExtensionTypes As Type() = Nothing
    '  Dim lobjActionExtensionType As Type = Nothing
    '  Dim lobjExtensionAssembly As Assembly = Nothing
    '  Dim lobjExtensionInterface As Type = Nothing
    '  Dim lobjExtensionInstance As Extensions.IActionExtension = Nothing

    '  Try

    '    lobjExtensionList = GetAvailableExtensions(lpFolderPath)
    '    'lobjOperationExtensionType = GetType(OperationExtension)
    '    For Each lstrExtension As String In lobjExtensionList
    '      lobjExtensionAssembly = Assembly.LoadFrom(lstrExtension)
    '      lobjExportedTypes = lobjExtensionAssembly.GetExportedTypes()
    '      For Each lobjExportedType As Type In lobjExportedTypes
    '        If lobjExportedType.IsAbstract Then
    '          Continue For
    '        End If
    '        lobjExtensionInterface = Nothing
    '        For Each lobjExportedInterface As Type In lobjExportedType.GetInterfaces
    '          If lobjExportedInterface.Name = "IActionExtension" Then
    '            lobjExtensionInstance = CType(lobjExtensionAssembly.CreateInstance(lobjExportedType.Name, True), Extensions.IActionExtension)
    '            If lobjExtensionInstance IsNot Nothing Then
    '              lobjExtensionActionList.Add(
    '                New ExtensionInformation(lobjExportedType.Name.Replace("Action", String.Empty),
    '                                         String.Empty, lobjExtensionInstance.CompanyName,
    '                                         lobjExtensionInstance.ProductName, lstrExtension))
    '            Else
    '              lobjExtensionActionList.Add(
    '                New ExtensionInformation(lobjExportedType.Name.Replace("Action", String.Empty),
    '                                         String.Empty, String.Empty, String.Empty, lstrExtension))
    '            End If
    '          End If
    '        Next
    '      Next
    '    Next

    '    Return lobjExtensionActionList

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    'Private Function GetAvailableExtensionActions() As IDictionary(Of String, Type)

    '  Dim lobjExtensionActionList As New SortedDictionary(Of String, Type)
    '  Dim lobjExtensionActions As IEnumerable(Of KeyValuePair(Of String, Type))
    '  Dim lobjActionType As Type = Nothing

    '  Try

    '    lobjActionType = GetType(IActionable)
    '    lobjExtensionActions = ExtensionCatalog.Instance.AvailableExtensions.Where(Function(t) lobjActionType.IsAssignableFrom(t.Value))

    '    For Each lobjExtensionPair As KeyValuePair(Of String, Type) In lobjExtensionActions
    '      lobjExtensionActionList.Add(lobjExtensionPair.Key, lobjExtensionPair.Value)
    '    Next

    '    Return lobjExtensionActionList

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try

    'End Function

    Private Sub RefreshCollections()
      Try
        mobjAvailableCoreActions = GetAvailableCoreActions()
        ' mobjAvailableExtensionOperations = GetAvailableExtensionOperations()
        mobjAvailableActions = GetAllAvailableActions()
      Catch Ex As Exception
        ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#Region "Create Methods"

    Private Function CreateAction(ByVal lpActionType As String) As IActionable

      Dim lobjActionType As Type = Nothing
      Dim lobjTypes As Object = Nothing
      Dim lobjRequestedActionType As Type = Nothing
      Dim lobjAction As IActionable = Nothing
      Dim lstrActionTypeName As String = lpActionType.Replace(" ", String.Empty)

      Try

        If AvailableActions.ContainsKey(lstrActionTypeName) = False Then
          Throw New UnknownActionException(lpActionType)

        Else

          Dim lobjCoreActionTypeDictionary As IDictionary(Of String, Type) = Instance.AvailableActions()

          If lobjCoreActionTypeDictionary IsNot Nothing AndAlso lobjCoreActionTypeDictionary.ContainsKey(lstrActionTypeName) Then

            lobjActionType = lobjCoreActionTypeDictionary.Item(lstrActionTypeName)
            lobjAction = CType(Activator.CreateInstance(lobjActionType), IActionable)

            Return lobjAction

          End If

          Throw New ArgumentException(String.Format("Unable to create core action of type '{0}': no action defined with that name.", lpActionType), "lpAction")
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw

      Finally
        ' Clean up references
        lobjActionType = Nothing
        lobjRequestedActionType = Nothing
      End Try

    End Function

    Private Function CreateAction(lpActionNode As XmlNode,
                                 lpLocale As String) As IActionable
      Try
        Dim lobjAction As IActionable
        Dim lstrActionName As String = Nothing
        Dim ldatStartTime As DateTime = DateTime.MinValue
        Dim ldatFinishTime As DateTime = DateTime.MinValue
        Dim lobjParameterNodes As XmlNodeList = Nothing
        Dim lobjParameter As IParameter = Nothing
        Dim lobjAttribute As XmlAttribute = Nothing
        Dim lobjNode As XmlNode = Nothing

        If lpActionNode Is Nothing Then
          Throw New ArgumentNullException("lpActionNode")
        End If

        If ((lpActionNode.Name.StartsWith("Run")) AndAlso (Not lpActionNode.Name.EndsWith("Action"))) Then
          Return GetRunOperable(lpActionNode)
        End If

        'If lpActionNode.HasChildNodes = False Then
        '  Throw New ArgumentException("Node has no child nodes", "lpActionNode")
        'End If

        ' Get the Name
        lobjAttribute = lpActionNode.Attributes("Name")

        If lobjAttribute Is Nothing Then
          'Throw New ArgumentException("Node has no name attribute")
          lstrActionName = lpActionNode.Name

        Else
          lstrActionName = lobjAttribute.InnerText
        End If

        ' Create the operation
        lobjAction = CType(TransformationActionFactory.Create(lstrActionName), IActionable)

        If lobjAction Is Nothing Then
          Throw New Exception(String.Format("Failed to create action of type '{0}'", lstrActionName))
        End If

        lobjParameterNodes = lpActionNode.SelectNodes("Parameters/*")

        If lobjParameterNodes IsNot Nothing Then

          For Each lobjParameterNode As XmlNode In lobjParameterNodes
            lobjParameter = Nothing
            lobjParameter = Parameter.Create(lobjParameterNode)

            If lobjParameter IsNot Nothing Then

              If lobjAction.Parameters.Contains(lobjParameter.Name) Then
                lobjAction.Parameters.Item(lobjParameter.Name) = lobjParameter

              Else
                lobjAction.Parameters.Add(lobjParameter)
              End If

            End If

          Next

          lobjAction.InitializeParameterValues()

        End If

        'If TypeOf lobjOperation Is IDecisionOperation Then

        '  ' TODO: Read the 'TrueOperations' node and the 'FalseOperations' node and update the respective collections.
        '  Dim lobjTrueOperationsNode As XmlNode = lpActionNode.SelectSingleNode("TrueOperations")

        '  If lobjTrueOperationsNode IsNot Nothing Then
        '    CType(lobjOperation, IDecisionOperation).TrueOperations.AddRange(Process.GetOperations(lobjTrueOperationsNode))
        '  End If

        '  Dim lobjFalseOperationsNode As XmlNode = lpActionNode.SelectSingleNode("FalseOperations")

        '  If lobjFalseOperationsNode IsNot Nothing Then
        '    CType(lobjOperation, IDecisionOperation).FalseOperations.AddRange(Process.GetOperations(lobjFalseOperationsNode))
        '  End If

        'End If

        Return lobjAction

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Private Function CreateAction(ByVal lpAction As IActionable) As IActionable

      Try

        ' This is used to clone an operation
        Dim lobjAction As IActionable

        ' Create the operation
        lobjAction = CType(TransformationActionFactory.Create(lpAction.Name), IActionable)

        ' Add the parameters
        lobjAction.Parameters = CType(lpAction.Parameters.Clone, IParameters)

        'If TypeOf lpOperation Is IDecisionOperation Then
        '  ' Clone the child operations as well
        '  CType(lobjOperation, IDecisionOperation).TrueOperations = CType(CType(lpOperation, IDecisionOperation).TrueOperations.Clone, IOperations)
        '  CType(lobjOperation, IDecisionOperation).FalseOperations = CType(CType(lpOperation, IDecisionOperation).FalseOperations.Clone, IOperations)
        'End If

        Return lobjAction

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

#End Region

#End Region

#Region "Event Handlers"

    'Private Sub mobjExtensionCatalog_CollectionChanged(sender As Object, e As Specialized.NotifyCollectionChangedEventArgs) Handles mobjExtensionCatalog.CollectionChanged
    '  Try
    '    RefreshCollections()
    '  Catch Ex As Exception
    '    ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

#End Region

  End Class

End Namespace
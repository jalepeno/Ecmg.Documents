'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  TransformationFactory.vb
'   Description :  [type_description_here]
'   Created     :  7/22/2013 9:48:31 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Reflection
Imports System.Xml
Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Utilities

#End Region

Namespace Transformations

  Public Class ActionFactory

#Region "Class Variables"

    Private Shared mintReferenceCount As Integer
    Private Shared mobjInstance As ActionFactory

    Private mobjAvailableActions As Generic.IDictionary(Of String, Type) = Nothing
    Private mobjAvailableActionDisplayNames As Generic.IList(Of String) = Nothing
    Private mobjAvailableCoreActions As Generic.IDictionary(Of String, Type) = Nothing

#End Region

#Region "Public Properties"

    Public ReadOnly Property AvailableActions() As Generic.IDictionary(Of String, Type)
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

    Public ReadOnly Property AvailableActionDisplayNames As Generic.IList(Of String)
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

    Public ReadOnly Property AvailableCoreActions As Generic.IDictionary(Of String, Type)
      Get
        Try
          If mobjAvailableCoreActions Is Nothing Then
            mobjAvailableCoreActions = GetAvailableCoreAction()
          End If

          Return mobjAvailableCoreActions

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Constructors"

    Private Sub New()
      Try
        mintReferenceCount = 0
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Singleton Support"

    Public Shared Function Instance() As ActionFactory

      Try

        If mobjInstance Is Nothing Then
          mobjInstance = New ActionFactory
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

#Region "Public Shared Methods"

    Public Shared Function Create(ByVal lpOperationType As String) As ITransformationAction

      Try
        Return Instance.CreateAction(lpOperationType)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Friend Shared Function Create(lpActionItem As IActionItem) As ITransformationAction
      Try
        Return Instance.CreateAction(lpActionItem)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Friend Shared Function Create(lpOperationNode As XmlNode,
                              lpLocale As String) As ITransformationAction
      Try
        Return Instance.CreateAction(lpOperationNode, lpLocale)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    Private Function GetAllAvailableActionDisplayNames() As Generic.IList(Of String)
      Try
        Dim lobjReturnList As New Generic.List(Of String)
        For Each lstrKey As String In AvailableActions.Keys
          lobjReturnList.Add(Helper.CreateDisplayName(lstrKey))
        Next

        Return lobjReturnList

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Private Function CreateActionItem(lpActionType As String) As IActionItem
    '  Try
    '    Beep()
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    Private Function CreateAction(lpActionNode As XmlNode,
                                 lpLocale As String) As ITransformationAction
      Try

        Dim lobjAction As ITransformationAction
        Dim lobjAttribute As XmlAttribute = Nothing
        Dim lstrName As String = Nothing
        Dim lobjParameterNodes As XmlNodeList = Nothing
        Dim lobjParameter As IParameter = Nothing

        If lpActionNode Is Nothing Then
          Throw New ArgumentNullException("lpActionNode")
        End If

        If lpActionNode.HasChildNodes = False Then
          Throw New ArgumentException("Node has no child nodes", "lpActionNode")
        End If

        ' Get the Name
        lobjAttribute = lpActionNode.Attributes("Name")

        If lobjAttribute Is Nothing Then
          'Throw New ArgumentException("Node has no name attribute")
          lstrName = lpActionNode.Name

        Else
          lstrName = lobjAttribute.InnerText
        End If

        ' Create the operation
        lobjAction = ActionFactory.Create(lstrName)

        If lobjAction Is Nothing Then
          Throw New Exception(String.Format("Failed to create action of type '{0}'", lstrName))
        End If

        lobjParameterNodes = lpActionNode.SelectNodes("Parameters/*")

        If lobjParameterNodes IsNot Nothing Then
          For Each lobjParameterNode As XmlNode In lobjParameterNodes
            lobjParameter = Nothing
            lobjParameter = Parameter.Create(lobjParameterNode)
            If lobjParameter IsNot Nothing Then
              If lobjParameter.Type = PropertyType.ecmEnum Then
                ' Make sure we get the standard values if they are not already set
                If lobjParameter.HasStandardValues = False Then
                  Dim lobjEnumParam As SingletonEnumParameter = lobjParameter

                  Dim lobjActionAssembly As Assembly
                  Try
                    lobjActionAssembly = Me.AvailableActions(lstrName).Assembly
                  Catch ex As Exception
                    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
                    lobjActionAssembly = Reflection.Assembly.GetExecutingAssembly
                  End Try

                  Dim lobjEnumType As System.Type = Helper.GetTypeFromAssembly(lobjActionAssembly, lobjEnumParam.EnumName)

                  If lobjEnumType IsNot Nothing Then
                    Dim lobjEnumDictionary As IDictionary(Of String, Integer) = Helper.EnumerationDictionary(lobjEnumType)
                    lobjEnumParam.SetStandardValues(lobjEnumDictionary.Keys)
                  End If
                End If
              End If

              If lobjAction.Parameters.Contains(lobjParameter.Name) Then
                lobjAction.Parameters.Item(lobjParameter.Name) = lobjParameter

              Else
                lobjAction.Parameters.Add(lobjParameter)
              End If

            End If
          Next
        End If

        Return lobjAction

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function CreateAction(ByVal lpActionType As String) As ITransformationAction

      Dim lobjActionType As Type = Nothing
      Dim lobjTypes As Object = Nothing
      Dim lobjRequestedOperationType As Type = Nothing
      Dim lobjAction As ITransformationAction = Nothing
      Dim lstrActionTypeName As String = lpActionType.Replace(" ", String.Empty)

      Try

        If AvailableActions.ContainsKey(lstrActionTypeName) = True Then
          ' We have it
        ElseIf AvailableActions.ContainsKey(String.Format("{0}Action", lstrActionTypeName)) = True Then
          lstrActionTypeName = String.Format("{0}Action", lstrActionTypeName)
        Else
          Throw New UnknownActionException(lstrActionTypeName)
        End If

        Dim lobjCoreActionTypeDictionary As Generic.IDictionary(Of String, Type) = Instance.AvailableActions()

        If lobjCoreActionTypeDictionary IsNot Nothing AndAlso lobjCoreActionTypeDictionary.ContainsKey(lstrActionTypeName) Then

          lobjActionType = lobjCoreActionTypeDictionary.Item(lstrActionTypeName)
          lobjAction = CType(Activator.CreateInstance(lobjActionType), ITransformationAction)
          lobjAction.Name = String.Empty

          Return lobjAction

        End If

        Throw New ArgumentException(String.Format("Unable to create core operation of type '{0}': no operation defined with that name.", lpActionType), "lpOperation")

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw

      Finally
        ' Clean up references
        lobjActionType = Nothing
        lobjRequestedOperationType = Nothing
      End Try
    End Function

    Private Function CreateAction(lpActionItem As IActionItem) As ITransformationAction

      Dim lobjTransformationAction As ITransformationAction = Nothing

      Try

        lobjTransformationAction = CreateAction(lpActionItem.Name)

        With lobjTransformationAction
          .Name = lpActionItem.Name
          .Description = lpActionItem.Description
        End With

        For Each lobjActionParameter As IParameter In lpActionItem.Parameters
          lobjTransformationAction.Parameters.SetItemByName(lobjActionParameter.Name, lobjActionParameter)
        Next

        lobjTransformationAction.InitializeParameterValues()

        Return lobjTransformationAction

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function GetAvailableCoreAction() As Generic.IDictionary(Of String, Type)

      Dim lobjOperationList As New Generic.SortedDictionary(Of String, Type)
      Dim lobjActionTypes As Generic.List(Of Type)
      Dim lobjAction As ITransformationAction = Nothing

      Try

        lobjActionTypes = GetAllCoreActionTypes()

        For Each lobjType As Type In lobjActionTypes

          lobjAction = CType(Activator.CreateInstance(lobjType), ITransformationAction)

          If lobjAction IsNot Nothing Then
            ' lobjOperationList.Add(lobjAction.Name, lobjType)
            lobjOperationList.Add(lobjType.Name, lobjType)
          End If

          lobjAction = Nothing

        Next

        Return lobjOperationList

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw

      Finally
        ' Clean up references
        lobjActionTypes = Nothing
      End Try

    End Function

    Private Function GetAllAvailableActions() As Generic.IDictionary(Of String, Type)
      Try

        Dim lobjAllOperationList As New Generic.SortedDictionary(Of String, Type)
        Dim lobjCoreOperationList As Generic.IDictionary(Of String, Type) = AvailableCoreActions
        ' Dim lobjExtensionOperationList As Generic.IDictionary(Of String, Type) = AvailableExtensionOperations

        For Each lstrCoreOperation As String In lobjCoreOperationList.Keys

          If lobjAllOperationList.ContainsKey(lstrCoreOperation) = False Then
            lobjAllOperationList.Add(lstrCoreOperation, lobjCoreOperationList.Item(lstrCoreOperation))
          End If

        Next

        'For Each lstrExtensionOperation As String In lobjExtensionOperationList.Keys

        '  If lobjAllOperationList.ContainsKey(lstrExtensionOperation) = False Then
        '    lobjAllOperationList.Add(lstrExtensionOperation, lobjExtensionOperationList.Item(lstrExtensionOperation))
        '  End If

        'Next

        Return lobjAllOperationList

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function GetAllCoreActionTypes() As Generic.List(Of Type)

      Dim lobjOperationList As New Generic.List(Of String)
      Dim lobjActionType As Type = Nothing
      Dim lobjAssembly As Reflection.Assembly = Nothing
      Dim lobjTypes As Generic.IEnumerable(Of Type) = Nothing
      Dim lobjCoreActionTypes As New Generic.List(Of Type)
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
            lobjCoreActionTypes.Add(lobjType)
          End If

          lobjAction = Nothing

        Next

        Return lobjCoreActionTypes

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

#End Region

  End Class

End Namespace
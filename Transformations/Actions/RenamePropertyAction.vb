'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Utilities

#End Region

Namespace Transformations

  ''' <summary>Action used to rename a property.</summary>
  <Serializable()>
  Public Class RenamePropertyAction
    Inherits PropertyAction
    'Implements IXmlSerializable

#Region "Class Constants"

    Private Const ACTION_NAME As String = "RenameProperty"
    Friend Const PARAM_NEW_NAME As String = "NewName"

#End Region

#Region "Class Variables"

    Private mstrNewName As String
    Private mstrNewSystemName As String

#End Region

#Region "Public Properties"

    Public Overrides ReadOnly Property ActionName As String
      Get
        Return ACTION_NAME
      End Get
    End Property

    Public Property NewName() As String
      Get
        Return mstrNewName
      End Get
      Set(ByVal Value As String)
        mstrNewName = Value
      End Set
    End Property

    Public Property NewSystemName() As String
      Get
        Return mstrNewSystemName
      End Get
      Set(ByVal Value As String)
        mstrNewSystemName = Value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(ActionType.RenameProperty)
    End Sub

    Public Sub New(ByVal lpPropertyName As String, ByVal lpNewName As String)
      MyBase.New(lpPropertyName, ActionType.RenameProperty)
      NewName = lpNewName
    End Sub

#End Region

#Region "Public Methods"

    Public Overrides Function Execute(ByRef lpErrorMessage As String) As ActionResult
      ' Rename a property
      Try

        If TypeOf Transformation.Target Is Document Then
          ' First do a pro-active check to see if the property exists
          If Transformation.Document.PropertyExists(PropertyScope, PropertyName, False) = False Then
            lpErrorMessage = String.Format("Unable to rename property '{0}' to '{1}'. The {2} '{0}' does not exist.",
                                 PropertyName, NewName, PropertyScope.ToString.ToLower.Replace("property", " property"))

            Return New ActionResult(Me, False, lpErrorMessage)
          End If

          If Transformation.Document.RenameProperty(PropertyScope, PropertyName, NewName, lpErrorMessage) = True Then
            Return New ActionResult(Me, True)
          Else
            Return New ActionResult(Me, False, String.Format("Failed to rename property '{0}' to '{1}': {2}",
                                                             PropertyName, NewName, lpErrorMessage))
          End If
        ElseIf TypeOf Transformation.Target Is Folder Then
          ' First do a pro-active check to see if the property exists
          If Transformation.Folder.PropertyExists(PropertyName, False) = False Then
            lpErrorMessage = String.Format("Unable to rename property '{0}' to '{1}'. The {2} '{0}' does not exist.",
                                 PropertyName, NewName, PropertyScope.ToString.ToLower.Replace("property", " property"))

            Return New ActionResult(Me, False, lpErrorMessage)
          End If

          If Transformation.Folder.RenameProperty(PropertyName, NewName, lpErrorMessage) = True Then
            Return New ActionResult(Me, True)
          Else
            Return New ActionResult(Me, False, String.Format("Failed to rename property '{0}' to '{1}': {2}",
                                                             PropertyName, NewName, lpErrorMessage))
          End If
        Else
          Throw New InvalidTransformationTargetException
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::Execute", Me.GetType.Name))
        If lpErrorMessage.Length = 0 Then
          lpErrorMessage = ex.Message
        End If
        Return New ActionResult(Me, False, lpErrorMessage)
      End Try

    End Function

#End Region

#Region "Protected Methods"

    Protected Overrides Function GetDefaultParameters() As IParameters
      Try
        Dim lobjParameters As IParameters = MyBase.GetDefaultParameters
        Dim lobjParameter As IParameter = Nothing


        If lobjParameters.Contains(PARAM_NEW_NAME) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_NEW_NAME,
            String.Empty, "The new name to be given to the property."))
        End If

        Return lobjParameters

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Friend Overrides Sub InitializeParameterValues()
      Try
        MyBase.InitializeParameterValues()
        Me.NewName = GetStringParameterValue(PARAM_NEW_NAME, String.Empty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace
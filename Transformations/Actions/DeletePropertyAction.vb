'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Utilities

Namespace Transformations
  ''' <summary>Action used to delete a property.</summary>
  <Serializable()>
  Public Class DeletePropertyAction
    Inherits PropertyAction

#Region "Class Constants"

    Private Const ACTION_NAME As String = "DeleteProperty"

#End Region

#Region "Public Properties"

    Public Overrides ReadOnly Property ActionName As String
      Get
        Return ACTION_NAME
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(ActionType.DeleteProperty)
    End Sub

    Public Sub New(ByVal lpPropertyName As String)
      MyBase.New(lpPropertyName, ActionType.DeleteProperty)
    End Sub

#End Region

#Region "Public Methods"

    Public Overrides Function Execute(ByRef lpErrorMessage As String) As ActionResult
      ' Delete a property
      Try

        If TypeOf Transformation.Target Is Document Then
          If Transformation.Document.PropertyExists(PropertyScope, PropertyName, False) Then
            Transformation.Document.DeleteProperty(PropertyScope, PropertyName)
            Return New ActionResult(Me, True)
          Else
            lpErrorMessage = String.Format("Unable to delete property '{0}', the property does not exist in the specified property scope of '{1}'.",
                                           PropertyName, PropertyScope.ToString)
            Return New ActionResult(Me, False, lpErrorMessage)
          End If
        ElseIf TypeOf Transformation.Target Is Folder Then
          If Transformation.Folder.PropertyExists(PropertyName, False) Then
            Transformation.Folder.DeleteProperty(PropertyName)
            Return New ActionResult(Me, True)
          Else
            lpErrorMessage = String.Format("Unable to delete property '{0}', the property does not exist in the specified property scope of '{1}'.",
                                           PropertyName, PropertyScope.ToString)
            Return New ActionResult(Me, False, lpErrorMessage)
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

  End Class

End Namespace
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  ReplaceFolderPath.vb
'   Description :  [type_description_here]
'   Created     :  6/19/2015 10:43:32 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Utilities

#End Region

Namespace Transformations

  Public Class ReplaceFolderPath
    Inherits Action

#Region "Class Constants"

    Private Const ACTION_NAME As String = "ReplaceFolderPath"
    Friend Const PARAM_OLD_PATH As String = "OldPath"
    Friend Const PARAM_NEW_PATH As String = "NewPath"

#End Region

#Region "Class Variables"

    Private mstrOldPath As String = String.Empty
    Private mstrNewPath As String = String.Empty

#End Region

#Region "Public Properties"

    Public Overrides ReadOnly Property ActionName As String
      Get
        Return ACTION_NAME
      End Get
    End Property

    Public Property OldPath() As String
      Get
        Return mstrOldPath
      End Get
      Set(ByVal Value As String)
        Try
          mstrOldPath = Value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property NewPath() As String
      Get
        Return mstrNewPath
      End Get
      Set(ByVal Value As String)
        Try
          mstrNewPath = Value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(lpOldPath As String, lpNewPath As String)
      Try
        mstrOldPath = lpOldPath
        mstrNewPath = lpNewPath
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Overrides Function Execute(ByRef lpErrorMessage As String) As ActionResult
      Try

        If TypeOf Transformation.Target Is Document Then
          Throw New InvalidTransformationTargetException()
        End If

        'If Transformation.Folder Is Nothing Then
        '  Throw New Exceptions.DocumentReferenceNotSetException
        'End If

        'Dim lobjFolderTree As New FolderTree(FolderPath)

        'lobjFolderTree.Folders

        If String.IsNullOrEmpty(OldPath) Then
          lpErrorMessage = "Unable to change folder path, the folder path value is not set."
          Return New ActionResult(Me, False, lpErrorMessage)
        Else
          ApplicationLogging.WriteLogEntry(String.Format("Replacing '{0}' with '{1}'.", OldPath, NewPath),
                                           Reflection.MethodBase.GetCurrentMethod(), TraceEventType.Information, 24387)
          Transformation.Folder.ReplacePath(OldPath, NewPath)
          Return New ActionResult(Me, True)
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
        Dim lobjParameters As IParameters = New Parameters
        Dim lobjParameter As IParameter = Nothing

        If lobjParameters.Contains(PARAM_OLD_PATH) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_OLD_PATH, String.Empty,
            "Specifies the old folder path to replace."))
        End If

        If lobjParameters.Contains(PARAM_NEW_PATH) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_NEW_PATH, String.Empty,
            "Specifies the new folder path to substitute."))
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
        Me.OldPath = GetStringParameterValue(PARAM_OLD_PATH, String.Empty)
        Me.NewPath = GetStringParameterValue(PARAM_NEW_PATH, String.Empty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace
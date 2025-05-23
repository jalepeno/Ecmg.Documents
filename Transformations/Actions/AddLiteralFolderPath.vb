﻿'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  AddLiteralFolderPath.vb
'   Description :  [type_description_here]
'   Created     :  5/20/2013 1:56:45 PM
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

  Public Class AddLiteralFolderPath
    Inherits Action

#Region "Class Constants"

    Private Const ACTION_NAME As String = "AddLiteralFolderPath"
    Friend Const PARAM_FOLDER_PATH As String = "FolderPath"
    Private Const DOCUMENT_CLASS_WILDCARD As String = "%DOCUMENT_CLASS%"

#End Region

#Region "Class Variables"

    Private mstrFolderPath As String = String.Empty

#End Region

#Region "Public Properties"

    Public Overrides ReadOnly Property ActionName As String
      Get
        Return ACTION_NAME
      End Get
    End Property

    Public Property FolderPath() As String
      Get
        Return mstrFolderPath
      End Get
      Set(ByVal Value As String)
        Try
          mstrFolderPath = Value
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

    Public Sub New(lpFolderPath As String)
      Try
        mstrFolderPath = lpFolderPath
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

        If TypeOf Transformation.Target Is Folder Then
          Throw New InvalidTransformationTargetException()
        End If

        If Transformation.Document Is Nothing Then
          Throw New Exceptions.DocumentReferenceNotSetException
        End If

        If String.IsNullOrEmpty(FolderPath) Then
          lpErrorMessage = "Unable to add folder path, the folder path value is not set."
          Return New ActionResult(Me, False, lpErrorMessage)
        Else
          ' Check to see if there is a document class wildcard
          If FolderPath.Contains(DOCUMENT_CLASS_WILDCARD) Then
            FolderPath = FolderPath.Replace(DOCUMENT_CLASS_WILDCARD, Transformation.Document.DocumentClass)
          End If
          Transformation.Document.AddFolderPath(FolderPath)
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

        If lobjParameters.Contains(PARAM_FOLDER_PATH) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_FOLDER_PATH, String.Empty,
            "Specifies the folder path to add to the document."))
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
        Me.FolderPath = GetStringParameterValue(PARAM_FOLDER_PATH, String.Empty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace
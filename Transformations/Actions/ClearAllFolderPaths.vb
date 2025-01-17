'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  ClearAllFolderPaths.vb
'   Description :  [type_description_here]
'   Created     :  5/20/2013 1:05:53 PM
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

  Public Class ClearAllFolderPaths
    Inherits Action

#Region "Class Constants"

    Private Const ACTION_NAME As String = "ClearAllFolderPaths"

#End Region

#Region "Public Properties"

    Public Overrides ReadOnly Property ActionName As String
      Get
        Return ACTION_NAME
      End Get
    End Property

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

        If Transformation.Document.FolderPathsProperty IsNot Nothing Then
          Transformation.Document.FolderPathsProperty.ClearPropertyValue()
        Else
          lpErrorMessage = String.Format("Unable to clear folder paths, there is no folder path property in document '{0}'.",
                               Transformation.Document.ID)

          Return New ActionResult(Me, False, lpErrorMessage)
        End If

        Return New ActionResult(Me, True)

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

    Protected Friend Overrides Sub InitializeParameterValues()
      Try
        ' Do nothing
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace
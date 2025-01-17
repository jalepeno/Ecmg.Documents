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

  ''' <summary>
  ''' Removes all versions without content.
  ''' </summary>
  ''' <remarks></remarks>
  <Serializable()>
  Public Class RemoveVersionsWithoutContentAction
    Inherits Action

#Region "Class Constants"

    Private Const ACTION_NAME As String = "RemoveVersionsWithoutContent"

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
      MyBase.New(ActionType.RemoveVersionsWithoutContentAction)
    End Sub

    Public Sub New(ByVal lpName As String)
      MyBase.New(ActionType.RemoveVersionsWithoutContentAction, lpName)
    End Sub

#End Region

#Region "Public Methods"

    Public Overrides Function Execute(ByRef lpErrorMessage As String) As ActionResult
      Try

        If TypeOf Transformation.Target Is Folder Then
          Throw New InvalidTransformationTargetException()
        End If

        Dim lobjActionResult As ActionResult = Nothing

        Dim lintRemovedVersionCount As Integer = Transformation.Document.RemoveVersionsWithoutContent

        Dim lstrDetails As String = String.Empty

        Select Case lintRemovedVersionCount
          Case 0
            lstrDetails = "No versions removed."
          Case 1
            lstrDetails = "1 version removed."
          Case Is > 1
            lstrDetails = String.Format("{0} versions removed.", lintRemovedVersionCount)
        End Select

        lobjActionResult = New ActionResult(Me, True, lstrDetails)

        Return lobjActionResult

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
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
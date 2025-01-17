'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

Namespace Transformations

  ''' <summary>
  ''' Used to stop the current transformation at the current step
  ''' </summary>
  ''' <remarks></remarks>
  <Serializable()>
  Public Class ExitTransformationAction
    Inherits Action

#Region "Class Constants"

    Private Const ACTION_NAME As String = "ExitTransformation"

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(ActionType.ExitTransformation)
    End Sub

    Public Sub New(ByVal lpName As String)
      MyBase.New(ActionType.ExitTransformation, lpName)
    End Sub

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

        Dim lobjActionResult As ActionResult = Nothing

        Me.Transformation.ShouldCancel = True

        lobjActionResult = New ActionResult(Me, True)

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
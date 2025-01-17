'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Core
Imports Documents.Utilities



Namespace Transformations
  ''' <summary>
  ''' Adds a content element to one or more versions of a document
  ''' </summary>
  ''' <remarks></remarks>
  <Serializable()>
  Public Class AddContentElement
    Inherits PropertyAction

#Region "Class Constants"

    Private Const ACTION_NAME As String = "AddContentElement"
    Friend Const PARAM_CONTENT_PATH As String = "ContentPath"
    Friend Const PARAM_CONTENT_ELEMENT_INDEX As String = "ContentElementIndex"

#End Region

#Region "Class Variables"

    Private mintContentElementIndex As Integer
    Private mstrContentPath As String = ""

#End Region

#Region "Public Properties"

    Public Overrides ReadOnly Property ActionName As String
      Get
        Return ACTION_NAME
      End Get
    End Property

    Public Property ContentElementIndex() As Integer
      Get
        Return mintContentElementIndex
      End Get
      Set(ByVal value As Integer)
        mintContentElementIndex = value
      End Set
    End Property

    Public Property ContentPath() As String
      Get
        Return mstrContentPath
      End Get
      Set(ByVal value As String)
        mstrContentPath = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(ActionType.ChangePropertyValue)
    End Sub


#End Region

#Region "Public Methods"

    Public Overrides Function Execute(ByRef lpErrorMessage As String) As ActionResult
      Try
        Throw New NotImplementedException
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Protected Methods"

    Protected Overrides Function GetDefaultParameters() As IParameters
      Try
        Dim lobjParameters As IParameters = New Parameters
        Dim lobjParameter As IParameter = Nothing

        lobjParameters.AddRange(MyBase.GetDefaultParameters)

        If lobjParameters.Contains(PARAM_CONTENT_PATH) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_CONTENT_PATH, String.Empty,
            "Specifies the path to the content to add."))
        End If

        If lobjParameters.Contains(PARAM_CONTENT_ELEMENT_INDEX) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_CONTENT_ELEMENT_INDEX, String.Empty,
            "Specifies where to insert the content element."))
        End If

        Return lobjParameters

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Public Overrides Sub InitializeFromParameters()
    '  Try

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    Protected Friend Overrides Sub InitializeParameterValues()
      Try
        MyBase.InitializeParameterValues()
        Me.ContentPath = GetStringParameterValue(PARAM_CONTENT_PATH, String.Empty)
        Me.ContentElementIndex = GetStringParameterValue(PARAM_CONTENT_ELEMENT_INDEX, String.Empty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace
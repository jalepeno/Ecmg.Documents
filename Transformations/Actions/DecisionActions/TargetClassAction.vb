'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Transformations

  ''' <summary>
  ''' An action that targets specific document classes and performs a set of transformation actions.
  ''' </summary>
  ''' <remarks></remarks>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class TargetClassAction
    Inherits Action
    Implements IDecisionAction

#Region "Class Constants"

    Private Const ACTION_NAME As String = "TargetClass"
    Friend Const PARAM_INCLUDE_RELATED_DOCUMENTS As String = "IncludeRelatedDocuments"

#End Region

#Region "Class Variables"

    Private mobjDocumentClassSet As New ClassActionsSets
    Private mblnIncludeRelatedDocuments As Boolean = True

#End Region

#Region "Public Properties"

    Public Overrides ReadOnly Property ActionName As String
      Get
        Return ACTION_NAME
      End Get
    End Property

    <XmlElement("DocumentClassActions")>
    Public Property DocumentClassSet As ClassActionsSets
      Get
        Return mobjDocumentClassSet
      End Get
      Set(value As ClassActionsSets)
        mobjDocumentClassSet = value
      End Set
    End Property

    ''' <summary>
    ''' Specifies whether or not to include related documents when evaluating 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlAttribute("IncludeRelatedDocuments")>
    Public Property IncludeRelatedDocuments As Boolean
      Get
        Return mblnIncludeRelatedDocuments
      End Get
      Set(value As Boolean)
        mblnIncludeRelatedDocuments = value
      End Set
    End Property

    ''' <summary>
    ''' Gets the action set associated with the current document class.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property CurrentActionSet As ClassActionSet
      Get
        Try
          Return GetCurrentClassSet()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(lpDocumentClassSet As ClassActionsSets)
      MyBase.New()
      Try
        DocumentClassSet = lpDocumentClassSet
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

        Dim lobjActionResult As ActionResult = Nothing

        ' We want this action to behave much like a Select Case statement.

        ' It will evaluate the current document class and if there is a named 
        ' class in our list then it will execute the specified actions for that class.  

        ' If there is no match for the current document class and there is a set of 
        ' default actions, it will perform those instead.

        ' If the flag 'IncludeRelatedDocuments' is set to true, all of the related documents 
        ' will be evaluated and if a match is found, they will be transformed as well.


        Dim lobjActionSet As ClassActionSet = CurrentActionSet

        If lobjActionSet IsNot Nothing Then
          ' Perform each of the actions in the set
          For Each lobjAction As Action In lobjActionSet.Actions
            lobjAction.Transformation = Me.Transformation
            lobjAction.Execute(lpErrorMessage)
          Next
        End If

        If IncludeRelatedDocuments AndAlso Me.Transformation.Document.Relationships.Count > 0 Then
          ' If we are supposed to include related documents, 
          ' we will pass this set of actions to each related document
          Dim lobjRelatedTransformation As New Transformation
          lobjRelatedTransformation.Name = "Related Document TargetClassAction Transformation"
          Dim lobjClonedAction As New TargetClassAction(Me.DocumentClassSet)
          lobjRelatedTransformation.Actions.Add(lobjClonedAction)

          For Each lobjRelationship As Relationship In Me.Transformation.Document.Relationships
            lobjRelationship.RelatedDocument.Transform(lobjRelatedTransformation, lpErrorMessage)
          Next
        End If

        ' If we are succesfull we need to return a successful result action.
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

    Protected Overrides Function GetDefaultParameters() As IParameters
      Try
        Dim lobjParameters As IParameters = New Parameters
        Dim lobjParameter As IParameter = Nothing

        If lobjParameters.Contains(PARAM_INCLUDE_RELATED_DOCUMENTS) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_INCLUDE_RELATED_DOCUMENTS, True,
            "Specifies whether or not to include related documents when evaluating."))
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
        Me.IncludeRelatedDocuments = GetBooleanParameterValue(PARAM_INCLUDE_RELATED_DOCUMENTS, True)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "IDecisionAction Implementation"

    Public Sub CombineActions() Implements IDecisionAction.CombineActions
      Try
        For Each lobjActionSet As ClassActionSet In DocumentClassSet
          lobjActionSet.Actions.Combine()
        Next
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public ReadOnly Property Evaluation As Boolean Implements IDecisionAction.Evaluation
      Get
        Try
          Return DocumentClassSet.ContainsClass(Transformation.CurrentDocumentClass)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    <XmlIgnore()>
    Public Property FalseActions As Actions Implements IDecisionAction.FalseActions
      Get
        Try
          Return DocumentClassSet.DefaultSet.Actions
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Actions)
        DocumentClassSet.DefaultSet.Actions = value
      End Set
    End Property

    <XmlIgnore()>
    Public ReadOnly Property RunActions As Actions Implements IDecisionAction.RunActions
      Get
        Try
          Return GetRunActions()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    <XmlIgnore()>
    Public Property TrueActions As Actions Implements IDecisionAction.TrueActions
      Get
        Try
          Return DocumentClassSet.Item(Transformation.CurrentDocumentClass).Actions
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Actions)
        Try
          Dim lobjActionSet As ClassActionSet = DocumentClassSet.Item(Transformation.CurrentDocumentClass)

          If lobjActionSet IsNot Nothing Then
            lobjActionSet.Actions = value
          End If

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

#End Region

#Region "Private Methods"

    Private Function GetCurrentClassSet() As ClassActionSet
      Try

        Dim lobjActionSet As ClassActionSet = DocumentClassSet.Item(Transformation.CurrentDocumentClass)

        If lobjActionSet IsNot Nothing Then
          Return lobjActionSet
        Else
          ' Get the default set
          Return DocumentClassSet.DefaultSet
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function GetRunActions() As Actions
      Try

        Return GetCurrentClassSet.Actions

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Friend Overloads Function DebuggerIdentifier() As String
      Try

        Dim lobjStringBuilder As New Text.StringBuilder

        lobjStringBuilder.AppendFormat("TargetClassAction: {0} Targeted Classes", DocumentClassSet.TargetClassCount)

        Select Case DocumentClassSet.DefaultSet.Actions.Count
          Case 0
            lobjStringBuilder.Append(" (No Default Actions)")
          Case 1
            lobjStringBuilder.AppendFormat(" ({0} Default Action)", DocumentClassSet.DefaultSet.Actions.Count)
          Case Else
            lobjStringBuilder.AppendFormat(" ({0} Default Actions)", DocumentClassSet.DefaultSet.Actions.Count)
        End Select

        Return lobjStringBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace

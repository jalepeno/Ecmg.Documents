'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Text
Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Transformations

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class DocumentClassTestAction
    Inherits Action
    Implements IDecisionAction

#Region "Class Constants"

    Private Const ACTION_NAME As String = "DocumentClassTest"
    Private Const PARAM_DOCUMENT_CLASS_TEST_VALUE As String = "DocumentClassTestValue"
    Private Const PARAM_CASE_SENSITIVE As String = "CaseSensitive"

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(ActionType.DocumentClassTestAction)
    End Sub

    Public Sub New(ByVal lpName As String)
      MyBase.New(ActionType.DocumentClassTestAction, lpName)
    End Sub

#End Region

#Region "Class Variables"

    Private mstrDocumentClass As String = String.Empty
    Private menuOperator As Data.Criterion.pmoStringComparisonOperator
    Private mblnCaseSensitive As Boolean = True
    Private mobjTrueActions As New Actions
    Private mobjFalseActions As New Actions
    Private mblnEvaluation As Nullable(Of Boolean)

#End Region

#Region "Public Properties"

    <XmlAttribute()>
    Public Overrides Property Name() As String
      Get
        Try

          If mstrName.Length = 0 Then

            Dim lobjStringBuilder As New StringBuilder

            If [Operator] = Data.Criterion.pmoStringComparisonOperator.opIsNullOrEmpty Then
              lobjStringBuilder.AppendFormat("{0}(DocumentClass {1}) {2} = {3}", Me.GetType.Name,
                                       Transformation.Document.DocumentClass, [Operator].ToString.Remove(0, 2), Evaluation)
            Else
              If CaseSensitive = True Then
                lobjStringBuilder.Append("Case Sensitive ")
              Else
                lobjStringBuilder.Append("Case Insensitive ")
              End If

              If Transformation.Document IsNot Nothing Then
                If String.IsNullOrEmpty(Transformation.Document.DocumentClass) Then
                  lobjStringBuilder.AppendFormat("{0}(DocumentClass {1} '{2}') = {3}", Me.GetType.Name,
                                           Transformation.Document.DocumentClass, [Operator].ToString.Remove(0, 2), DocumentClassTestValue, Evaluation)
                Else
                  lobjStringBuilder.AppendFormat("{0}('{1}' {2} '{3}') = {4}", Me.GetType.Name,
                                           Transformation.Document.DocumentClass, [Operator].ToString.Remove(0, 2), DocumentClassTestValue, Evaluation)
                End If
              End If
            End If

            If Evaluation = True Then
              lobjStringBuilder.AppendFormat(" with {0} Actions", TrueActions.Count)
            Else
              lobjStringBuilder.AppendFormat(" with {0} Actions", FalseActions.Count)
            End If

            mstrName = lobjStringBuilder.ToString

          End If

          Return mstrName

        Catch ex As Exception
          Return MyBase.mstrName
        End Try
      End Get
      Set(ByVal value As String)
        mstrName = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the Document Class name value to be tested against.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Xml.Serialization.XmlAttribute("TestValue")>
    Public Property DocumentClassTestValue() As String
      Get
        Return mstrDocumentClass
      End Get
      Set(ByVal value As String)
        mstrDocumentClass = value
      End Set
    End Property

    <Xml.Serialization.XmlAttribute("Operator")>
    Public Property [Operator]() As Data.Criterion.pmoStringComparisonOperator
      Get
        Return menuOperator
      End Get
      Set(ByVal value As Data.Criterion.pmoStringComparisonOperator)
        menuOperator = value
      End Set
    End Property

    <Xml.Serialization.XmlAttribute()>
    Public Property CaseSensitive() As Boolean
      Get
        Return mblnCaseSensitive
      End Get
      Set(ByVal value As Boolean)
        mblnCaseSensitive = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the actions to be executed if the decision evaluates to True
    ''' </summary>
    Public Property TrueActions() As Actions Implements IDecisionAction.TrueActions
      Get
        Return mobjTrueActions
      End Get
      Set(ByVal value As Actions)
        mobjTrueActions = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the actions to be executed if the decision evaluates to False
    ''' </summary>
    Public Property FalseActions() As Actions Implements IDecisionAction.FalseActions
      Get
        Return mobjFalseActions
      End Get
      Set(ByVal value As Actions)
        mobjFalseActions = value
      End Set
    End Property

    Public ReadOnly Property Evaluation As Boolean Implements IDecisionAction.Evaluation
      Get
        Try
          If mblnEvaluation.HasValue = False Then
            mblnEvaluation = EvaluateDocumentClass()
          End If

          If mblnEvaluation.HasValue Then
            Return mblnEvaluation
          Else
            Return False
          End If

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ''' <summary>
    ''' Gets the set of actions to be run based on the current evaluation.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property RunActions As Actions Implements IDecisionAction.RunActions
      Get
        Try
          If Evaluation = True Then
            Return TrueActions
          Else
            If mblnEvaluation.HasValue Then
              Return FalseActions
            Else
              Return New Actions
            End If
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Public Methods"

    Public Overrides Function Execute(ByRef lpErrorMessage As String) As ActionResult
      Try

        Dim lobjActionResult As ActionResult = Nothing
        Dim lstrDetails As String = String.Empty
        Dim lstrSourceDocumentClass As String = Transformation.Document.DocumentClass

        If String.IsNullOrEmpty(lstrSourceDocumentClass) Then
          Return New ActionResult(Me, False, "Document Class is not set in source document.")
        End If

        If Evaluation = True Then
          For Each lobjTrueAction As Action In TrueActions
            lobjTrueAction.Transformation = Me.Transformation
            lobjTrueAction.Execute(lpErrorMessage)
          Next
        ElseIf Evaluation = False Then
          For Each lobjFalseAction As Action In FalseActions
            lobjFalseAction.Transformation = Me.Transformation
            lobjFalseAction.Execute(lpErrorMessage)
          Next
        End If

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

    Public Sub CombineActions() Implements IDecisionAction.CombineActions
      Try
        TrueActions.Combine()
        FalseActions.Combine()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Protected Methods"

    Protected Overrides Function GetDefaultParameters() As IParameters
      Try
        Dim lobjParameters As IParameters = New Parameters
        Dim lobjParameter As IParameter = Nothing

        If lobjParameters.Contains(PARAM_DOCUMENT_CLASS_TEST_VALUE) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_DOCUMENT_CLASS_TEST_VALUE, String.Empty,
            "Specifies the document class to test for."))
        End If

        If lobjParameters.Contains(PARAM_CASE_SENSITIVE) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_CASE_SENSITIVE, True,
            "Specifies where to insert the content element."))
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
        Me.DocumentClassTestValue = GetStringParameterValue(PARAM_DOCUMENT_CLASS_TEST_VALUE, String.Empty)
        Me.CaseSensitive = GetBooleanParameterValue(PARAM_CASE_SENSITIVE, True)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

    Private Shadows Function DebuggerIdentifier() As String
      Try
        ' Return Name

        Dim lobjReturnBuilder As New StringBuilder

        If Not String.IsNullOrEmpty(Name) Then
          lobjReturnBuilder.AppendFormat("{0} - ", Name)
        End If

        lobjReturnBuilder.AppendFormat("{0}: {1}({4}{2}{4}) is {3}", Me.GetType.Name,
                                       [Operator].ToString.Substring(2), DocumentClassTestValue,
                                       Evaluation, ControlChars.Quote)

        If RunActions.Count = 1 Then
          lobjReturnBuilder.AppendFormat(" - {0}", RunActions.Item(0).DebuggerIdentifier)
        Else
          lobjReturnBuilder.AppendFormat(" - {0} actions", RunActions.Count)
        End If

        Return lobjReturnBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function EvaluateDocumentClass() As Nullable(Of Boolean)
      Try

        If Transformation.Document Is Nothing Then
          Return Nothing
        End If

        Dim lstrSourceDocumentClass As String = Transformation.Document.DocumentClass

        If [Operator] = Data.Criterion.pmoStringComparisonOperator.opIsNullOrEmpty Then
          If String.IsNullOrEmpty(lstrSourceDocumentClass) Then
            Return True
          Else
            Return False
          End If
        End If

        If String.IsNullOrEmpty(lstrSourceDocumentClass) Then
          Return False
        End If

        ' First do any of the tests that are based on String.Compare 
        ' since it can take a case sensitivity argument.
        Select Case [Operator]
          Case Data.Criterion.pmoStringComparisonOperator.opEquals
            If String.Compare(lstrSourceDocumentClass, DocumentClassTestValue, Not CaseSensitive) = 0 Then
              Return True
            Else
              Return False
            End If

          Case Data.Criterion.pmoStringComparisonOperator.opNotEqual
            If String.Compare(lstrSourceDocumentClass, DocumentClassTestValue, Not CaseSensitive) <> 0 Then
              Return True
            Else
              Return False
            End If

        End Select

        ' If the test is not to be case sensitive 
        ' then set both values to the same case.
        If CaseSensitive = False Then
          lstrSourceDocumentClass = lstrSourceDocumentClass.ToLower
          DocumentClassTestValue = DocumentClassTestValue.ToLower
        End If

        Select Case [Operator]
          Case Data.Criterion.pmoStringComparisonOperator.opStartsWith
            If lstrSourceDocumentClass.StartsWith(DocumentClassTestValue) Then
              Return True
            Else
              Return False
            End If

          Case Data.Criterion.pmoStringComparisonOperator.opEndsWith
            If lstrSourceDocumentClass.EndsWith(DocumentClassTestValue) Then
              Return True
            Else
              Return False
            End If

          Case Data.Criterion.pmoStringComparisonOperator.opContains
            If lstrSourceDocumentClass.Contains(DocumentClassTestValue) Then
              Return True
            Else
              Return False
            End If

          Case Data.Criterion.pmoStringComparisonOperator.opIn
            If DocumentClassTestValue.Contains(lstrSourceDocumentClass) Then
              Return True
            Else
              Return False
            End If

        End Select

        ' If we have made it this far then we could not make a decision.
        ApplicationLogging.WriteLogEntry("Unable to evaluate document class", Reflection.MethodBase.GetCurrentMethod, TraceEventType.Warning, 65243)
        Return New Nullable(Of Boolean)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return False
      End Try
    End Function

#End Region

  End Class

End Namespace
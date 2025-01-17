'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Transformations

  ''' <summary>Based on a defined evaluation, runs one or more transformation actions</summary>
  ''' <remarks>TrueActions run if True, FalseActions run if False</remarks>
  <Serializable()>
  Public Class DecisionAction
    Inherits ChangePropertyValue
    Implements IDecisionAction

#Region "Class Variables"

    Private mobjTrueActions As New Actions
    Private mobjFalseActions As New Actions
    Private mblnEvaluation As Nullable(Of Boolean)
    Private mstrTestValue As String
    Private mstrTestValueDataType As Core.PropertyType = PropertyType.ecmString
    Private menuOperator As Data.Criterion.pmoOperator
    Private mblnCaseSensitive As Boolean = True

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the value to be tested against
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TestValue() As String
      Get
        Return mstrTestValue
      End Get
      Set(ByVal value As String)
        mstrTestValue = value
      End Set
    End Property

    Public Property TestValueType() As PropertyType
      Get
        Return mstrTestValueDataType
      End Get
      Set(ByVal value As PropertyType)
        mstrTestValueDataType = value
      End Set
    End Property

    <Xml.Serialization.XmlAttribute("Operator")>
    Public Property [Operator]() As Data.Criterion.pmoOperator
      Get
        Return menuOperator
      End Get
      Set(ByVal value As Data.Criterion.pmoOperator)
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

    Public ReadOnly Property RunActions As Actions Implements IDecisionAction.RunActions
      Get
        Try
          If mblnEvaluation.HasValue = False Then
            Return New Actions
          Else
            If Evaluation = True Then
              Return TrueActions
            Else
              Return FalseActions
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

#Region "Private Properties"

    ''' <summary>
    ''' Gets or sets the evaluation value for the decision
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected ReadOnly Property Evaluation() As Boolean Implements IDecisionAction.Evaluation
      Get
        Return mblnEvaluation
      End Get
    End Property

#End Region

#Region "Costructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpName As String,
                   ByVal lpPropertyName As String,
                   Optional ByVal lpPropertyScope As Core.PropertyScope = Core.PropertyScope.VersionProperty,
                   Optional ByVal lpVersionIndex As Integer = Transformation.TRANSFORM_ALL_VERSIONS,
                   Optional ByVal lpDataLookup As DataLookup = Nothing,
                   Optional ByVal lpTestValue As String = "",
                   Optional ByVal lpTestValueType As Core.PropertyType = PropertyType.ecmString,
                   Optional ByVal lpOperator As Data.Criterion.pmoOperator = Data.Criterion.pmoOperator.opEquals,
                   Optional ByVal lpCaseSensitive As Boolean = True)
      MyBase.New(lpPropertyName, lpPropertyScope, lpVersionIndex, ValueSource.DataLookup, lpDataLookup)
      Try
        Name = lpName
        TestValue = lpTestValue
        TestValueType = lpTestValueType
        [Operator] = lpOperator
        CaseSensitive = lpCaseSensitive
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Overrides Function Execute(ByRef lpErrorMessage As String) As ActionResult
      Try

        If SourceExists() = False Then
          ' We were not able to verify the source property above
          If TypeOf (DataLookup) Is IPropertyLookup Then
            lpErrorMessage = String.Format("Unable to make decision, the source property {1} does not exist.",
                                           PropertyName, CType(DataLookup, IPropertyLookup).SourceProperty.PropertyName)
          Else
            lpErrorMessage = "Unable to make decision, the source does not exist."
          End If
          Return New ActionResult(Me, False, lpErrorMessage)
        End If

        Dim lobjActionResult As ActionResult = Nothing

        'Test for true or false
        '  TODO: Determine how to define evaluation statement

        ' Make an evaluation for true or false
        Select Case SourceType
          Case ChangePropertyValue.ValueSource.Literal

            ' We will not process Literals here
            lpErrorMessage += "SourceType of 'Literal' is not supported for DecisionAction, use 'DataLookup' or 'DataMap' instead."
            ApplicationLogging.WriteLogEntry(lpErrorMessage, TraceEventType.Error)

            Return New ActionResult(Me, False, lpErrorMessage)

          Case ChangePropertyValue.ValueSource.DataLookup
            Try
              Dim lobjValue As Object
              If Me.VersionIndex = Transformation.TRANSFORM_ALL_VERSIONS Then
                For lintVersionCounter As Integer = 0 To Transformation.Document.Versions.Count - 1
                  Dim lobjVersion As Version = Transformation.Document.Versions(lintVersionCounter)
                  lobjValue = DataLookup.GetValue(lobjVersion)
                  mblnEvaluation = Evaluate(lobjValue)
                  If Evaluation = True Then
                    Exit For
                  End If
                Next
              Else
                lobjValue = DataLookup.GetValue(Transformation.Document)
                mblnEvaluation = Evaluate(lobjValue)
              End If
              ' Return New ActionResult(Me, True)
            Catch ex As Exception
              ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
              If ex.Message.StartsWith("No value found for the expression") = True Then
                'Throw New Exception("DataMap found no value for [" & lobjChangePropertyValueAction.PropertyName & "]", ex)
                lpErrorMessage &= "DataMap found no value for [" & PropertyName & "]" & ex.Message
                Return New ActionResult(Me, False, lpErrorMessage)
              Else
                'Throw New Exception("DataLookup Failed", ex)
                lpErrorMessage &= "DataLookup Failed; " & Helper.FormatCallStack(ex)
                Return New ActionResult(Me, False, lpErrorMessage)
              End If
            Finally

            End Try

        End Select

        Dim lintActionCounter As Integer

        If Evaluation = True Then
          If TrueActions.Count > 0 Then
            'Transformation.'LogSession.LogDebug("Decision Action '{0}' evaluated to True.  About to execute {1} True actions.", Me.Name, TrueActions.Count)
          Else
            'Transformation.'LogSession.LogDebug("Decision Action '{0}' evaluated to True.  No True actions found.", Me.Name)
          End If

          For Each lobjTrueAction As Action In TrueActions
            lintActionCounter += 1
            lobjTrueAction.Transformation = Me.Transformation
            'Transformation.'LogSession.LogDebug("About to execute transformation True action '{0}: {1}'.", lintActionCounter, lobjTrueAction.Name)
            'Transformation.'LogSession.LogString(Level.Debug, "Action Details", lobjTrueAction.ToXmlString())
            lobjTrueAction.Execute(lpErrorMessage)
          Next
        Else
          If FalseActions.Count > 0 Then
            'Transformation.'LogSession.LogDebug("Decision Action '{0}' evaluated to False.  About to execute {1} False actions.", Me.Name, FalseActions.Count)
          Else
            'Transformation.'LogSession.LogDebug("Decision Action '{0}' evaluated to False.  No False actions found.", Me.Name)
          End If

          For Each lobjFalseAction As Action In FalseActions
            lintActionCounter += 1
            lobjFalseAction.Transformation = Me.Transformation
            'Transformation.'LogSession.LogDebug("About to execute transformation False action '{0}: {1}'.", lintActionCounter, lobjFalseAction.Name)
            'Transformation.'LogSession.LogString(Level.Debug, "Action Details", lobjFalseAction.ToXmlString())
            lobjFalseAction.Execute(lpErrorMessage)
          Next
        End If

        lobjActionResult = New ActionResult(Me, True, lpErrorMessage)

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

#Region "Private Methods"

    ''' <summary>
    ''' Compares the supplied value with the DecisionAction TestValue
    ''' </summary>
    ''' <param name="lpValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function Evaluate(ByVal lpValue As Object) As Boolean
      Try

        If lpValue Is Nothing Then
          'Transformation.'LogSession.LogDebug("DecisionAction.Evaluate: Value to check is Null")
        End If

        Dim lobjTestValue As Object = TestValue

        If lobjTestValue Is Nothing Then
          'Transformation.'LogSession.LogDebug("DecisionAction.Evaluate: TestValue is Null")
        End If

        ' Try to coerce the test value to the specified data type
        Select Case TestValueType
          Case PropertyType.ecmDate
            lobjTestValue = CDate(TestValue)
          Case PropertyType.ecmBoolean
            lobjTestValue = CBool(TestValue)
          Case PropertyType.ecmLong
            lobjTestValue = CLng(TestValue)
          Case PropertyType.ecmDouble
            lobjTestValue = CDbl(TestValue)
        End Select

        If lobjTestValue IsNot Nothing Then
          'Transformation.'LogSession.LogDebug("DecisionAction.Evaluate: TestValue = '{0}'", lobjTestValue.ToString())
        End If

        If lpValue IsNot Nothing Then
          'Transformation.'LogSession.LogDebug("DecisionAction.Evaluate: Value to check = '{0}'", lpValue.ToString())
        End If

        Select Case [Operator]
          Case Data.Criterion.pmoOperator.opEquals
            ' If lobjTestValue = lpValue Then
            If CaseSensitive = True Then
              If String.Equals(lobjTestValue, lpValue, StringComparison.InvariantCulture) Then
                Return True
              Else
                Return False
              End If
            Else
              If String.Equals(lobjTestValue, lpValue, StringComparison.InvariantCultureIgnoreCase) Then
                Return True
              Else
                Return False
              End If
            End If

          Case Data.Criterion.pmoOperator.opGreaterThan
            If lobjTestValue > lpValue Then
              Return True
            Else
              Return False
            End If

          Case Data.Criterion.pmoOperator.opGreaterThanOrEqualTo
            If lobjTestValue >= lpValue Then
              Return True
            Else
              Return False
            End If

          Case Data.Criterion.pmoOperator.opIn
            ' We will only support the In operator for strings
            If TestValueType = PropertyType.ecmString _
            AndAlso CStr(lobjTestValue).Contains(lpValue.ToString()) Then
              Return True
            Else
              Return False
            End If

          Case Data.Criterion.pmoOperator.opLessThan
            If lobjTestValue < lpValue Then
              Return True
            Else
              Return False
            End If

          Case Data.Criterion.pmoOperator.opLessThanOrEqualTo
            If lobjTestValue <= lpValue Then
              Return True
            Else
              Return False
            End If

          Case Data.Criterion.pmoOperator.opLike
            If TestValueType <> PropertyType.ecmString Then
              Return False
            End If

            If TestValue.StartsWith("*"c) Then
              If CStr(lpValue).EndsWith(TestValue.Substring(1)) Then
                Return True
              End If
            End If

            If TestValue.EndsWith("*"c) Then
              If CStr(lpValue).StartsWith(TestValue) Then
                Return True
              End If
            End If

            If lpValue.ToString.Contains(TestValue) Then
              Return True
            End If

        End Select
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace

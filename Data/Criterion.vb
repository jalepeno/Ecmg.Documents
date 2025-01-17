'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------


Imports Documents.Core
Imports Documents.Utilities

Namespace Data

  ' <summary>Fully describes a single criterion for use in a search.</summary>
  'When I run content loader on the vista 64-bit machine i have to use <Serializable()> or else I get a datacontract error.  Need to try on 32-bit and see what happens
  <Serializable()>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class Criterion
    'Inherits Search.Criterion
    Implements IEquatable(Of Criterion)

#Region "Enumerations"

    Public Enum pmoDataType
      ecmBinary = 1
      ecmBoolean = 2
      ecmDate = 3
      ecmDouble = 4
      ecmGuid = 5
      ecmLong = 6
      ecmObject = 7
      ecmString = 8
    End Enum

    Public Enum pmoOperator
      opEquals = 0
      opLessThan = 1
      opGreaterThan = 2
      opLessThanOrEqualTo = 3
      opGreaterThanOrEqualTo = 4
      opLike = 5 'Example would be %bob%
      opNotLike = 6
      opIn = 7
      opEndsWith = 8 'Similar to "like" but does this %bob
      opContentContains = 9 'Content search, like Verity
      opStartsWith = 10 'Similar to "like" but does this bob%
      opNotEqual = 11
    End Enum

    Public Enum pmoStringComparisonOperator
      opEquals = 0
      opNotEqual = 1
      opIsNullOrEmpty = 2
      opStartsWith = 3
      opEndsWith = 4
      opContains = 5
      opIn = 6
    End Enum

#End Region

#Region "Class Variables"

    Private mstrName As String
    Private mstrDocumentPropertyName As String
    'Private mstrVersionPropertyName As String
    Private menuOperator As pmoOperator
    Private menuPropertyScope As Core.PropertyScope
    Private menuDataType As pmoDataType = pmoDataType.ecmString
    Private menuSetEvaluation As SetEvaluation = SetEvaluation.seAnd
    Private mstrValue As String
    Private mobjValues As New Core.Values
    Private menuCardinality As Core.Cardinality = Core.Cardinality.ecmSingleValued

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' The descriptive name for the criterion. The value is optional.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Name() As String
      Get
        Return mstrName
      End Get
      Set(ByVal Value As String)
        mstrName = Value
      End Set
    End Property

    ''' <summary>
    ''' The property name as defined in the repository.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PropertyName() As String
      Get
        Return mstrDocumentPropertyName
      End Get
      Set(ByVal Value As String)
        mstrDocumentPropertyName = Value
      End Set
    End Property

    ''' <summary>
    ''' The scope of the property to include in the search criteria. Possible values are 'VersionProperty' and 'DocumentProperty'.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PropertyScope() As Core.PropertyScope
      Get
        Return menuPropertyScope
      End Get
      Set(ByVal Value As Core.PropertyScope)
        menuPropertyScope = Value
      End Set
    End Property

    ''' <summary>
    ''' The operator to evaluate the property against. Possible values are opEquals, opLessThan, opGreaterThan, opLessThanOrEqualTo, opGreaterThanOrEqualTo, opLike, opNotLike, opIn.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property [Operator]() As pmoOperator
      Get
        Return menuOperator
      End Get
      Set(ByVal Value As pmoOperator)
        menuOperator = Value
      End Set
    End Property

    ''' <summary>
    ''' Determines how two or more Criterion objects will relate to one another in a search. Possible values are 'seAnd' and 'seOr'.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SetEvaluation() As SetEvaluation
      Get
        Return menuSetEvaluation
      End Get
      Set(ByVal Value As SetEvaluation)
        menuSetEvaluation = Value
      End Set
    End Property

    ''' <summary>
    ''' The data type the criteria will be evaluated as
    ''' </summary>
    ''' <value></value>
    ''' <returns>A pmoDataType enumeration value</returns>
    ''' <remarks>If not set, will default to ecmString</remarks>
    Public Property DataType() As pmoDataType
      Get
        Return menuDataType
      End Get
      Set(ByVal Value As pmoDataType)
        menuDataType = Value
      End Set
    End Property

    Public ReadOnly Property WhereClause() As String
      Get
        Try
          Return GetWhereClause()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ''' <summary>
    ''' Single-valued property value. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Not used when searching based on document values.</remarks>
    Public Property Value() As String
      Get
        Return mstrValue
      End Get
      Set(ByVal value As String)
        mstrValue = value
      End Set
    End Property

    ''' <summary>
    ''' Collection of multi-valued property values. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Not used when searching based on document values.</remarks>
    Public Property Values() As Core.Values
      Get
        Return mobjValues
      End Get
      Set(ByVal value As Core.Values)
        mobjValues = value
      End Set
    End Property

    ''' <summary>
    ''' Specifies whether or not the specified property is a 
    ''' single valued or multi valued field. Possible values 
    ''' are 'ecmSingleValued' and 'ecmMultiValued'.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Cardinality() As Core.Cardinality
      Get
        Return menuCardinality
      End Get
      Set(ByVal value As Core.Cardinality)
        menuCardinality = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpName As String,
                   Optional ByVal lpPropertyName As String = Nothing,
                   Optional ByVal lpPropertyScope As Core.PropertyScope = Core.PropertyScope.VersionProperty,
                   Optional ByVal lpOperator As pmoOperator = pmoOperator.opEquals,
                   Optional ByVal lpSetEvaluation As SetEvaluation = SetEvaluation.seAnd) ', ByVal lpDataType As pmoDataType)
      Try
        Name = lpName
        PropertyName = lpPropertyName
        PropertyScope = lpPropertyScope
        'VersionPropertyName = lpVersionPropertyName
        [Operator] = lpOperator
        SetEvaluation = lpSetEvaluation
        'DataType = lpDataType
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

    Protected Friend Overridable Function DebuggerIdentifier() As String
      Try
        If Cardinality = Core.Cardinality.ecmSingleValued Then
          If DataType = pmoDataType.ecmString Then
            Return String.Format("{0}'{1}'", WhereClause, Me.Value).Replace("''", "'")
          Else
            Return String.Format("{0}{1}", WhereClause, Me.Value)
          End If
        Else
          Return String.Format("{0}{1}", WhereClause, Me.Values.GetFirstValue)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function GetWhereClause() As String
      Try
        Dim lstrWhereClause As String = String.Empty ' = String.Format("[{0}]", Name)

        Select Case [Operator]
          Case pmoOperator.opEquals
            'lstrWhereClause &= " = "
            lstrWhereClause = String.Format("[{0}] = ", Name)

          Case pmoOperator.opGreaterThan
            'lstrWhereClause &= " > "
            lstrWhereClause = String.Format("[{0}] > ", Name)

          Case pmoOperator.opGreaterThanOrEqualTo
            'lstrWhereClause &= " >= "
            lstrWhereClause = String.Format("[{0}] >= ", Name)

          Case pmoOperator.opIn
            'lstrWhereClause &= " IN ("
            lstrWhereClause = String.Format("[{0}] IN (", Name)
          Case pmoOperator.opLessThan
            'lstrWhereClause &= " < "
            lstrWhereClause = String.Format("[{0}] < ", Name)

          Case pmoOperator.opLessThanOrEqualTo
            'lstrWhereClause &= " <= "
            lstrWhereClause = String.Format("[{0}] <= ", Name)

          Case pmoOperator.opLike, pmoOperator.opEndsWith
            'lstrWhereClause &= " LIKE '%"
            lstrWhereClause = String.Format("[{0}] LIKE '%", Name)

          Case pmoOperator.opStartsWith
            'lstrWhereClause &= " LIKE '"
            lstrWhereClause = String.Format("[{0}] LIKE '", Name)

          Case pmoOperator.opNotLike
            'lstrWhereClause &= " NOT LIKE '%"
            lstrWhereClause = String.Format("[{0}] NOT LIKE '%", Name)

          Case pmoOperator.opContentContains
            'lstrWhereClause = String.Empty
            'lstrWhereClause &= " CONTAINS(Content,'"
            lstrWhereClause = " CONTAINS(Content,'"

          Case pmoOperator.opNotEqual
            'lstrWhereClause &= " <> "
            lstrWhereClause = String.Format("[{0}] <> ", Name)

        End Select

        Return lstrWhereClause
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Public Operator Overloads"

    Public Shared Operator =(ByVal c1 As Criterion, ByVal c2 As Criterion) As Boolean
      Try
        ' Compare all the parts of the criterion
        Return c1.Cardinality = c2.Cardinality AndAlso
               c1.Name = c2.Name AndAlso
               c1.Operator = c2.Operator AndAlso
               c1.PropertyName = c2.PropertyName AndAlso
               c1.PropertyScope = c2.PropertyScope AndAlso
               c1.SetEvaluation = c2.SetEvaluation AndAlso
               c1.Value = c2.Value AndAlso c1.Values.ToString = c2.Values.ToString
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Operator

    Public Shared Operator <>(ByVal c1 As Criterion, ByVal c2 As Criterion) As Boolean
      Try
        Return Not (c1 = c2)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Operator

    Public Overloads Function Equals(ByVal c As Criterion) As Boolean Implements IEquatable(Of Data.Criterion).Equals
      Try
        If c Is Nothing Then
          Return False
        Else
          Return c = Me
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overloads Shared Function Equals(ByVal c1 As Criterion, ByVal c2 As Criterion) As Boolean
      Try
        Return (c1 = c2)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
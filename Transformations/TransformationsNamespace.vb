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

''' <summary>
''' Contains objects used for transforming documents
''' </summary>
Namespace Transformations

#Region "Public Enumerations"

  ''' <summary>
  ''' This enumeration has been deprecated
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum ActionType
    CreateProperty = 0
    RenameProperty = 1
    DeleteProperty = 2
    ChangePropertyValue = 3
    DeletePropertiesNotInDocumentClass = 4
    ChangePropertyCardinality = 5
    ChangeAllTimesToUTC = 6
    ExitTransformation = 7
    RunTransformation = 8
    DecisionAction = 9
    AddContentElement = 10
    AddPropertyValue = 11
    ChangeContentMimeType = 12
    ChangeContentRetrievalName = 13
    ChangeDateTimePropertyValueToUTC = 14
    RemoveVersionsWithoutContentAction = 15
    DocumentClassTestAction = 16
    ClearAllValues = 17
    'TrimPropertyValue = 4
  End Enum

  ''' <summary>
  ''' Specifies which versions are to be affected by an operation
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum VersionScope
    AllVersions = 0
    FirstVersionOnly = 1
    LastVersionOnly = 2
  End Enum

  Public Enum TrimPropertyType
    Left = 0
    Right = 1
    Both = 2
  End Enum

  Public Enum LookupType
    Map = 0
    Parser = 1
    Script = 2
    List = 3
  End Enum

#End Region

#Region "Transformation Interfaces"

  Public Interface IPropertyLookup
    Property SourceProperty() As LookupProperty
    Property DestinationProperty() As LookupProperty

    Function GetParameters() As IParameters

  End Interface

  Public Interface ITransformation
    Inherits IDescription

    Property Actions() As ITransformationActions
    ReadOnly Property DisplayName As String
    Property Document() As Document
    ReadOnly Property TransformationFilePath() As String

    Function TransformDocument(ByVal lpEcmDocument As Document,
                                      Optional ByRef lpErrorMessage As String = "") As Document
  End Interface

  Public Interface ITransformationAction
    Inherits IDescription
    ReadOnly Property DisplayName As String
    Property Parameters As IParameters
    Property Transformation() As ITransformation

    Sub InitializeParameterValues()

  End Interface

  Public Interface ITransformationActions
    Inherits ICollection(Of ITransformationAction)

  End Interface

  Public Interface IDecisionAction
    ReadOnly Property Evaluation As Boolean
    Property TrueActions As Actions
    Property FalseActions As Actions
    ReadOnly Property RunActions As Actions
    ''' <summary>
    ''' Looks through all actions in both the FalseActions and 
    ''' TrueActions collections and if a RunTransformationAction 
    ''' is present, replaces it with all the actions in the target 
    ''' transformation.
    ''' </summary>
    ''' <remarks></remarks>
    Sub CombineActions()
  End Interface

#End Region

#Region "Exclusion Objects"

  '  Public Class Exclusions
  '    Inherits Core.CCollection

  '#Region "Class Variables"

  '#End Region

  '#Region "Public Properties"

  '#End Region

  '#Region "Constructors"

  '    Public Sub New()

  '    End Sub

  '#End Region

  '#Region "Overloads"

  '#Region "Item Overloads"

  '    Default Shadows Property Item(ByVal Index As Integer) As Exclusion
  '      Get
  '        Return MyBase.Item(Index)
  '      End Get
  '      Set(ByVal value As Exclusion)
  '        MyBase.Item(Index) = value
  '      End Set
  '    End Property

  '    Default Shadows Property Item(ByVal Name As String) As Exclusion
  '      Get
  '        Dim lobjExclusion As Exclusion
  '        For lintCounter As Integer = 0 To MyBase.Count - 1
  '          lobjExclusion = CType(MyBase.Item(lintCounter), Exclusion)
  '          If lobjExclusion.Name = Name Then
  '            Return lobjExclusion
  '          End If
  '        Next
  '        Throw New Exception("There is no Item by the Name '" & Name & "'.")
  '        'Throw New InvalidArgumentException
  '      End Get
  '      Set(ByVal value As Exclusion)
  '        Dim lobjExclusion As Exclusion
  '        For lintCounter As Integer = 0 To MyBase.Count - 1
  '          lobjExclusion = CType(MyBase.Item(lintCounter), Exclusion)
  '          If lobjExclusion.Name = Name Then
  '            MyBase.Item(lintCounter) = value
  '          End If
  '        Next
  '        Throw New Exception("There is no Item by the Name '" & Name & "'.")
  '      End Set
  '    End Property

  '#End Region

  '    Public Overloads Sub Add(ByVal lpExclusion As Exclusion)
  '      MyBase.Add(lpExclusion)
  '    End Sub

  '    Public Overloads Function Delete(ByVal ID As String) As Boolean
  '      Try
  '        Delete(Item(ID))
  '        Return True
  '      Catch ex As Exception
  '        Return False
  '      End Try
  '    End Function

  '    Public Overloads Sub Sort()
  '      MyBase.Sort()
  '    End Sub

  '#End Region

  '  End Class

#End Region

#Region "Results Objects"

  ''' <summary>Detailed result returned from an Action.Execute method.</summary>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class ActionResult
    Inherits Core.ActionResult

#Region "Class Variables"

    Private mobjAction As Action

#End Region

#Region "Public Properties"

    Public Property Action() As Action
      Get
        Return mobjAction
      End Get
      Set(ByVal value As Action)
        mobjAction = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpAction As Action, ByVal lpSuccess As Boolean, Optional ByVal lpDetails As String = "")
      MyBase.New(lpSuccess, lpDetails)
      Try
        Action = lpAction
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Overrides Function ToResultCode() As Result
      Try
        Return MyBase.ToResultCode()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    Friend Overrides Function DebuggerIdentifier() As String
      Try
        If Action IsNot Nothing AndAlso Not String.IsNullOrEmpty(Action.Name) Then
          Return String.Format("{0} {1}", Action.Name, MyBase.DebuggerIdentifier)
        Else
          Return MyBase.DebuggerIdentifier
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

#End Region

End Namespace
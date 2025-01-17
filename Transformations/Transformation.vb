'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.SerializationUtilities
Imports Documents.Utilities

#End Region

Namespace Transformations

  ''' <summary>Worker object used to transform document metadata.</summary>
  ''' <remarks>Serialized instances should use the file extension .ctf (Content Transformation File)</remarks>
  <Serializable()>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class Transformation
    Implements ISerialize
    Implements ICloneable
    Implements IDescription
    Implements IDisposable
    Implements ITransformation
    Implements ITransformationConfiguration
    'Implements IXmlSerializable
    'Implements ILoggable

#Region "Class Constants"

    ''' <summary>
    ''' Constant integer value used to specify that all versions of a document should be
    ''' transofrmed.
    ''' </summary>
    Public Const TRANSFORM_ALL_VERSIONS As Integer = -1

    ''' <summary>
    ''' Constant value representing the file extension to save XML serialized instances
    ''' to.
    ''' </summary>
    Public Const TRANSFORMATION_FILE_EXTENSION As String = "ctf"

    Private Const XML_HEADER As String = "<?xml version=""1.0"" encoding=""utf-16""?>"

#End Region

#Region "Class Variables"

    Private mstrName As String = String.Empty
    Private mstrDescription As String = String.Empty
    Private mstrDisplayName As String = String.Empty
    Private mstrExclusionPath As String = String.Empty
    'Private mobjExclusions As New Exclusions
    Private mobjActions As New Actions(Me)
    Private mblnShouldCancel As Boolean
    Private mstrTransformationFilePath As String = String.Empty
    Private mobjDocument As Document
    Private mobjFolder As Folder
    Private mobjTarget As IMetaHolder

#End Region

#Region "Public Properties"

    Public ReadOnly Property DisplayName As String Implements ITransformation.DisplayName, ITransformationConfiguration.DisplayName ', IProcess.DisplayName
      Get
        Try
          If String.IsNullOrEmpty(mstrDisplayName) Then
            mstrDisplayName = Helper.CreateDisplayName(Me.Name)
          End If
          Return mstrDisplayName
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      'Set(value As String)
      '  mstrDisplayName = value
      'End Set
    End Property

    ''' <summary>The name of the transformation.</summary>
    <XmlAttribute()>
    Public Property Name() As String Implements IDescription.Name, INamedItem.Name, ITransformationConfiguration.Name
      Get
        Return mstrName
      End Get
      Set(ByVal value As String)
        mstrName = value
      End Set
    End Property

    ''' <summary>A description for the transformation.</summary>
    Public Property Description() As String Implements IDescription.Description, ITransformationConfiguration.Description
      Get
        Return mstrDescription
      End Get
      Set(ByVal value As String)
        mstrDescription = value
      End Set
    End Property

    '''' <summary>The directory path to copy excluded documents to.</summary>
    'Public Property ExclusionPath() As String
    '  Get
    '    Try
    '      If Right(mstrExclusionPath, 1) <> "\" Then
    '        Return mstrExclusionPath & "\"
    '      Else
    '        Return mstrExclusionPath
    '      End If
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, "Transformation::Get_ExclusionPath")
    '      Return ""
    '    End Try
    '  End Get
    '  Set(ByVal Value As String)
    '    mstrExclusionPath = Value
    '  End Set
    'End Property

    'Public ReadOnly Property ExclusionFolderPath(ByVal lpExclusion As Exclusion) As String
    '  Get

    '    Dim lstrFolderPath As String = ""

    '    Try

    '      lstrFolderPath = ExclusionPath & lpExclusion.FolderName

    '      If Right(lstrFolderPath, 1) <> "\" Then
    '        lstrFolderPath &= "\"
    '      End If

    '      Try
    '        If Not Directory.Exists(lstrFolderPath) Then
    '          Directory.CreateDirectory(lstrFolderPath)
    '        End If
    '      Catch ex As Exception
    '        ApplicationLogging.LogException(ex, String.Format("Transformation::Get_ExclusionPath_CreateDirectory('{0}')", lstrFolderPath))
    '      End Try

    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, "Transformation::Get_ExclusionPath")
    '    End Try

    '    Return lstrFolderPath

    '  End Get
    'End Property

    '''' <summary>The collection of Exclusion objects.</summary>
    'Public Property Exclusions() As Exclusions
    '  Get
    '    Return mobjExclusions
    '  End Get
    '  Set(ByVal Value As Exclusions)
    '    mobjExclusions = Value
    '  End Set
    'End Property

    ''' <summary>The collection of transformation actions.</summary>
    Public Property Actions() As Actions
      Get
        Return mobjActions
      End Get
      Set(ByVal Value As Actions)
        mobjActions = Value
      End Set
    End Property

    <XmlIgnore()>
    Public Property ActionItems As IActionItems Implements ITransformationConfiguration.Actions
      Get
        Try
          ' <Modified by: Ernie at 9/22/2014-2:12:43 PM on machine: ERNIE-THINK>
          'Return mobjActions
          ' Actions and IActionItems are not currently directly transferable.  
          ' Until they are we will stop returning anything for this property.
          Return New ActionItems
          ' </Modified by: Ernie at 9/22/2014-2:12:43 PM on machine: ERNIE-THINK>
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As IActionItems)
        Try
          ' <Modified by: Ernie at 9/22/2014-2:14:06 PM on machine: ERNIE-THINK>
          'mobjActions = value
          ' Actions and IActionItems are not currently directly transferable.  
          ' Until they are we will stop returning anything for this property.
          ' </Modified by: Ernie at 9/22/2014-2:14:06 PM on machine: ERNIE-THINK>
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    <XmlIgnore()>
    Public Property Target As IMetaHolder
      Get
        Try
          Return mobjTarget
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As IMetaHolder)
        Try
          mobjTarget = value
          If value IsNot Nothing Then
            If TypeOf (mobjTarget) Is Document Then
              mobjDocument = mobjTarget
            ElseIf TypeOf (mobjTarget) Is Folder Then
              mobjFolder = mobjTarget
            End If
          Else
            mobjDocument = Nothing
            mobjFolder = Nothing
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>The document to transform.</summary>
    <XmlIgnore()>
    Public Property Document() As Document Implements ITransformation.Document
      Get
        Try
          If (Helper.IsDeserializationBasedCall = False) AndAlso (Helper.IsSerializationBasedCall = False) Then
            If mobjDocument IsNot Nothing Then
              Return mobjDocument
            Else
              If Helper.CallStackContainsMethodName("InvokeMethod") Then
                ' If this was called as part of a reflection based discovery go ahead and return nothing.
                Return mobjDocument
              Else
                ' Make sure the caller knows that the document reference has not been set.
                Throw New Exceptions.TransformationDocumentReferenceNotSetException(Me)
              End If
            End If
          Else
            Return Nothing
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try

      End Get
      Set(ByVal value As Document)
        Try
          mobjTarget = value
          mobjDocument = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    <XmlIgnore()>
    Public Property Folder As Folder
      Get
        Try
          Return mobjFolder
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Folder)
        Try
          mobjTarget = value
          mobjFolder = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public ReadOnly Property CurrentDocumentClass As String
      Get
        Try
          If mobjDocument IsNot Nothing Then
            Return Document.DocumentClass
          Else
            Return String.Empty
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ''' <summary>
    ''' Gets the file location the transformation object was constructed from if available.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property TransformationFilePath() As String Implements ITransformation.TransformationFilePath
      Get
        Return mstrTransformationFilePath
      End Get
    End Property



#End Region

#Region "Protected Properties"

    Protected Friend Property ShouldCancel() As Boolean
      Get
        Return mblnShouldCancel
      End Get
      Set(ByVal value As Boolean)
        mblnShouldCancel = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Private Sub LoadFromXmlDocument(ByVal lpXML As Xml.XmlDocument)
      Try
        Dim lobjTransformation As Transformation = DeSerialize(lpXML)
        Helper.AssignObjectProperties(lobjTransformation, Me)
        'With Me
        '  .ExclusionPath = lobjTransformation.ExclusionPath
        '  .Exclusions = lobjTransformation.Exclusions
        '  .Actions = lobjTransformation.Actions
        'End With
      Catch ex As Exception
        ApplicationLogging.LogException(ex, "Transformation::LoadFromXmlDocument(lpXml)")
      End Try
    End Sub

    ''' <summary>Default Constructor</summary>
    Public Sub New()

    End Sub

    ''' <summary>
    ''' Constructs a Transformation object using the specified Actions, Exclusions and
    ''' ExclusionPath.
    ''' </summary>
    ''' <param name="lpActions">The collection of Actions to assign to the new Transformation.</param>
    ''' <param name="lpExclusions">The collection of Exclusions to assign to the new Transformation.</param>
    ''' <param name="lpExclusionPath">The path to save excluded files to.</param>
    Public Sub New(ByVal lpActions As Actions)

      Try
        Actions = lpActions
      Catch ex As Exception
        ApplicationLogging.LogException(ex, "Transformation::New(lpActions)")
      End Try

    End Sub

    ''' <summary>
    ''' Constructs a Transformation object by deserializing the XML in the specified
    ''' path.
    ''' </summary>
    ''' <param name="lpXMLFilePath">A fully qualified XML file path for the serialized Transformation file.</param>
    Public Sub New(ByVal lpXMLFilePath As String)

      Dim lobjXMLDocument As New Xml.XmlDocument
      Dim lstrLogMessage As String = String.Empty

      Try
        lobjXMLDocument.Load(lpXMLFilePath)
        mstrTransformationFilePath = lpXMLFilePath
        LoadFromXmlDocument(lobjXMLDocument)


        ' If the transformation name was not specified 
        ' then use the file name as the transformation name.
        If Me.Name.Length = 0 Then
          Me.Name = Path.GetFileNameWithoutExtension(lpXMLFilePath)
        End If

        ' Update any relative paths if necessary
        'ReconcileRelativeTransformationPaths()

        ' Write an entry in the log
        Select Case Actions.Count
          Case 0
            lstrLogMessage = String.Format("Created Transformation '{1}' with no actions using file '{0}'",
                                           lpXMLFilePath, Me.Name)
            ApplicationLogging.WriteLogEntry(lstrLogMessage, TraceEventType.Warning, 4504)
          Case 1
            lstrLogMessage = String.Format("Successfully created Transformation '{1}' with 1 action using file '{0}'",
                                           lpXMLFilePath, Me.Name)
            ApplicationLogging.WriteLogEntry(lstrLogMessage, TraceEventType.Information, 4500)
          Case Else
            lstrLogMessage = String.Format("Successfully created Transformation '{1}' with {2} actions using file '{0}'",
                                           lpXMLFilePath, Me.Name, Me.Actions.Count)
            ApplicationLogging.WriteLogEntry(lstrLogMessage, TraceEventType.Information, 4500)
        End Select

        ' For some reason we are seeing cases where property values are being deserialized as XmlNode arrays.
        ' The following method cleans those up.
        CleanUpXmlNodeArrays()

        ' Ensure that all actions have a valid reference to the parent transformation
        For Each lobjAction As Action In Me.Actions
          lobjAction.Transformation = Me
        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("Transformation::New(lpXMLFilePath:'{0}')", lpXMLFilePath))
        Throw New ApplicationException(String.Format("Unable to create transformation object from xml file: {0}", ex.Message), ex)
      End Try
    End Sub

    Public Sub New(lpStream As Stream)
      Try
        Dim lobjTransformation As Transformation = DeSerialize(lpStream)
        Helper.AssignObjectProperties(lobjTransformation, Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpXML As Xml.XmlDocument)
      Try
        LoadFromXmlDocument(lpXML)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(lpTransformationConfiguration As ITransformationConfiguration)
      Try

        Dim lobjNewAction As ITransformationAction

        With Me
          .Name = lpTransformationConfiguration.Name
          .Description = lpTransformationConfiguration.Description
        End With

        For Each lobjAction As IActionItem In lpTransformationConfiguration.Actions
          lobjNewAction = ActionFactory.Create(lobjAction)
          If lobjNewAction IsNot Nothing Then
            Actions.Add(lobjNewAction)
          End If
        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Combines all the actions from any RunTransformation actions into the parent Transformation.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Combine() As Transformation
      Try

        Dim lobjNewTransformation As New Transformation

        ' Set the new name from the original
        lobjNewTransformation.Name = Me.Name

        ' Build the description from the original
        lobjNewTransformation.Description = String.Format("Combined from {0}", Name)
        If Not String.IsNullOrEmpty(Me.Description) Then
          lobjNewTransformation.Description = String.Format("{0} - {1}", lobjNewTransformation.Description, Me.Description)
        End If

        ' Set the parent transformation file path
        lobjNewTransformation.mstrTransformationFilePath = Me.TransformationFilePath

        For Each lobjAction As Action In Actions
          If TypeOf lobjAction Is IDecisionAction Then
            DirectCast(lobjAction, IDecisionAction).CombineActions()
          End If
          If TypeOf lobjAction Is RunTransformationAction Then
            Dim lobjRunActionTarget As Transformation = DirectCast(lobjAction, RunTransformationAction).FindTargetTransformation
            If lobjRunActionTarget IsNot Nothing Then
              lobjNewTransformation.Actions.Add(lobjRunActionTarget.Combine.Actions)
            End If
          Else
            lobjNewTransformation.Actions.Add(lobjAction)
          End If
        Next

        Return lobjNewTransformation

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Gets all the child transformations.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetChildTransformations() As IList(Of Transformation)
      Try
        Return Actions.GetChildTransformations
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>Deletes all document properties that do not include values.</summary>
    ''' <remarks>Used to simplify the and streamline the document properties.</remarks>
    Public Shared Function DeletePropertiesWithoutValues(ByVal lpEcmDocument As Core.Document,
                                                  Optional ByRef lpErrorMessage As String = "") As Core.Document


      Return lpEcmDocument.DeletePropertiesWithoutValues(lpErrorMessage)

      ''Dim lobjProperty As ECMProperty

      'Try
      '  ' Go through all the document properties
      '  lpEcmDocument.Properties = DeletePropertiesWithoutValues(lpEcmDocument.Properties, lpErrorMessage)

      '  ' Go though all of the version properties
      '  For Each lobjVersion As Version In lpEcmDocument.Versions
      '    lobjVersion.Properties = DeletePropertiesWithoutValues(lobjVersion.Properties, lpErrorMessage)
      '  Next

      '  Return lpEcmDocument

      'Catch ex As Exception
      '  lpErrorMessage = Ecmg.ECMUtilities.Helper.FormatCallStack(ex)
      '  Return Nothing
      'End Try

    End Function

    ''' <summary>
    ''' Checks all the document properties for the declared data type. If the data type
    ''' is different from the data type specified in the destination repository, the document
    ''' property data type is updated to reflect the data type defined in the destination
    ''' repository.
    ''' </summary>
    Public Shared Function SetPropertyTypesToDestinationTypes(ByVal lpEcmDocument As Core.Document,
                                                         ByVal lpDestinationProperties As Core.ClassificationProperties, Optional ByRef lpErrorMessage As String = "") As Core.Document

      Try

        lpEcmDocument.Properties = SetPropertyTypesToDestinationTypes(lpEcmDocument.Properties, lpDestinationProperties, lpErrorMessage)

        For Each lobjVersion As Core.Version In lpEcmDocument.Versions
          lobjVersion.Properties = SetPropertyTypesToDestinationTypes(lobjVersion.Properties, lpDestinationProperties, lpErrorMessage)
        Next

        Return lpEcmDocument

      Catch ex As Exception
        ApplicationLogging.LogException(ex, "Transformation::SetPropertyTypesToDestinationTypes(lpEcmDocument, lpDestinationProperties, lpErrorMessage")
        lpErrorMessage = Helper.FormatCallStack(ex)
        Return lpEcmDocument
      End Try


    End Function

    ''' <summary>
    ''' Checks all the document properties for the declared data type. If the data type
    ''' is different from the data type specified in the destination repository, the document
    ''' property data type is updated to reflect the data type defined in the destination
    ''' repository.
    ''' </summary>
    Public Shared Function SetPropertyTypesToDestinationTypes(ByRef lpECMProperties As Core.ECMProperties,
                                                       ByVal lpDestinationProperties As Core.ClassificationProperties, Optional ByRef lpErrorMessage As String = "") As Core.ECMProperties

      Dim lobjDestinationProperty As ECMProperty = Nothing
      Dim lobjReplacementProperty As ECMProperty = Nothing
      Dim lobjECMProperty As ECMProperty = Nothing

      Try
        'For Each lobjECMProperty As Core.ECMProperty In lpECMProperties
        For lintPropertyCounter As Integer = lpECMProperties.Count - 1 To 0 Step -1
          Try

            lobjECMProperty = lpECMProperties(lintPropertyCounter)

            ' Try to get the destination property by the same name
            lobjDestinationProperty = lpDestinationProperties(lobjECMProperty.Name)

            ' We we found the property then get the type
            If lobjDestinationProperty IsNot Nothing AndAlso
              lobjECMProperty.Type <> lobjDestinationProperty.Type Then

              ' We will make an exception for the case where the property name is 
              ' 'Folders Filed In' and the source type is string and the destination type is object.
              ' We want to always pass 'Folders Filed In' as a collection of folder path strings.
              If String.Equals(lobjDestinationProperty.Name, Document.FOLDER_PATHS_PROPERTY_NAME) Then
                If (lobjECMProperty.Type = PropertyType.ecmString) AndAlso
                  (lobjDestinationProperty.Type = PropertyType.ecmObject) Then
                  Continue For
                End If
              End If

              ' In order to change the type we will need to recreate the property
              ' since the different property types have different classes associated with them.
              lobjReplacementProperty = PropertyFactory.Create(lobjDestinationProperty.Type, lobjDestinationProperty.Name,
                                                               lobjDestinationProperty.SystemName, lobjDestinationProperty.Cardinality)

              ' Set the property value or values
              If lobjECMProperty.HasValue Then
                If lobjReplacementProperty.Cardinality = Cardinality.ecmSingleValued Then
                  lobjReplacementProperty.Value = lobjECMProperty.Value
                Else
                  lobjReplacementProperty.Values = lobjECMProperty.Values
                End If
              End If

              ' Replace the property
              lpECMProperties.Replace(lobjECMProperty.Name, lobjReplacementProperty)

              ' Write an entry in the log
              ApplicationLogging.WriteLogEntry(String.Format(
                 "Changed the data type for property '{0}' to '{1}' in transformation action SetPropertyTypesToDestinationTypes",
                 lobjECMProperty.Name, lobjECMProperty.Type.ToString), TraceEventType.Information, 2010)
            End If

          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            ' Skip it and move on
          End Try
        Next
      Catch ex As Exception
        ApplicationLogging.LogException(ex, "Transformation::SetPropertyTypesToDestinationTypes(lpECMProperties, lpDestinationProperties, lpErrorMessage")
        lpErrorMessage = Helper.FormatCallStack(ex)
        Return lpECMProperties
      End Try

      Return lpECMProperties


    End Function

    ''' <summary>
    ''' Transforms the specified document using the set of transformation actions
    ''' specified for the instance.
    ''' </summary>
    Public Function TransformDocument(ByVal lpEcmDocument As Core.Document,
                                      Optional ByRef lpErrorMessage As String = "") As Core.Document _
                                    Implements ITransformation.TransformDocument

      ''LogSession.EnterMethod(Level.Debug, Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))
      ''LogSession.LogDebug("Transforming document '{0}' with {1} actions", lpEcmDocument.ID, Me.Actions.Count)
      ApplicationLogging.WriteLogEntry(String.Format("Transforming document '{0}' with {1} actions", lpEcmDocument.ID, Me.Actions.Count), TraceEventType.Verbose)
      ''LogSession.LogString(Level.Debug, "Transformation Details", Me.ToXmlString())

      If lpEcmDocument Is Nothing Then
        Throw New ArgumentNullException("lpEcmDocument", "lpEcmDocument argument is null")
      End If

      ' Can't log the document xml here.  It closes the content stream.
      ' Ernie Bahr 07/22/2021
      ''LogSession.LogString(Level.Debug, "Pre-Transformation Document Details", lpEcmDocument.ToXmlString())

      Me.Document = lpEcmDocument
      Dim lstrTransformationError As String = String.Empty
      Dim lobjActionResult As Transformations.ActionResult
      Dim lintActionCounter As Integer
      Dim lintErrorCounter As Integer
      Dim lstrLogMessage As String = Nothing

      Try

        ' In case we cancelled the previous transformation as a result of an 
        ' ExitTransformationAction we need to reset the ShouldCancel flag.  
        ' Otherwise, none of the actions will get executed.
        ShouldCancel = False

        For Each lobjAction As Action In Actions
          lstrTransformationError = String.Empty
          lintActionCounter += 1
          If Not lobjAction Is Nothing Then

            lobjAction.Transformation = Me

            ' Check to see if we have been asked to cancel out
            If ShouldCancel = True Then
              Exit For
            End If

            ''LogSession.LogDebug("About to execute transformation action '{0}: {1}'.", lintActionCounter, lobjAction.Name)
            ApplicationLogging.WriteLogEntry(String.Format("About to execute transformation action '{0}: {1}'.", lintActionCounter, lobjAction.Name), TraceEventType.Verbose)
            ''LogSession.LogObject(Level.Debug, "Action Details", lobjAction)
            ''LogSession.LogString(Level.Debug, "Action Details", lobjAction.ToXmlString())

            lobjActionResult = lobjAction.Execute(lstrTransformationError)
            If lobjActionResult.Success = False Then
              lintErrorCounter += 1
              ' Make a notation in the log
              lstrLogMessage = String.Format("Transformation action {0}. ({1}) (Document Id: {3}) failed: {2}",
                                             lintActionCounter.ToString, lobjAction.Name, lstrTransformationError, Me.Document.ID)
              ApplicationLogging.WriteLogEntry(lstrLogMessage, TraceEventType.Warning, 32852)
              If lpErrorMessage.Length = 0 Then
                lpErrorMessage &= lstrTransformationError
              Else
                lpErrorMessage &= "; " & lstrTransformationError
              End If
            Else
              If lobjActionResult.Details.Length > 0 Then
                'If lpErrorMessage.Length = 0 Then
                '  lpErrorMessage &= lobjActionResult.Details
                'Else
                '  lpErrorMessage &= "; " & lobjActionResult.Details
                'End If
              End If
              If lstrTransformationError IsNot Nothing AndAlso lstrTransformationError.Length > 0 Then
                lstrLogMessage = String.Format("Transformation action {0}. ({1}) completed with the message: {2}",
                                             lintActionCounter.ToString, lobjAction.Name, lstrTransformationError)
                ApplicationLogging.WriteLogEntry(lstrLogMessage, TraceEventType.Warning, 32853)
              End If
            End If
          End If
        Next

        '  Delete all of the non-persistent properties created during the transformation
        lpEcmDocument.DeleteTemporaryProperties()

        lpEcmDocument.Versions.Sort()
        lpEcmDocument.Versions.SortProperties()

        ' Update the header
        ' First check to see if we have a header
        If lpEcmDocument.Header Is Nothing Then
          lpEcmDocument.UpdateHeader()
        End If

        ' Add the transformation information to the header
        lpEcmDocument.Header.TransformationSeries.Add(New Document.Modification(Environment.MachineName, Environment.UserName, Me))

        ' Add an entry to the log
        Select Case lintErrorCounter
          Case 0
            lstrLogMessage = String.Format("Transformed document '{0}' using transform '{1}' with {2} actions.",
                                           lpEcmDocument.ID, Me.Name, Actions.Count)
            ApplicationLogging.WriteLogEntry(lstrLogMessage, TraceEventType.Information, 6767)
          Case 1
            lstrLogMessage = String.Format("Transformed document '{0}' using transform '{1}' with {2} failure out of {3} actions",
                                           lpEcmDocument.ID, Me.Name, lintErrorCounter, Actions.Count)
            ApplicationLogging.WriteLogEntry(lstrLogMessage, TraceEventType.Warning, 6768)
          Case Else
            lstrLogMessage = String.Format("Transformed document '{0}' using transform '{1}' with {2} failures out of {3} actions",
                                           lpEcmDocument.ID, Me.Name, lintErrorCounter, Actions.Count)
            ApplicationLogging.WriteLogEntry(lstrLogMessage, TraceEventType.Warning, 6768)
        End Select

        ' Can't log the document xml here.  It closes the content stream.
        ' Ernie Bahr 07/22/2021
        ' 'LogSession.LogString(Level.Debug, "Post-Transformation Document Details", lpEcmDocument.ToXmlString())

        Return lpEcmDocument

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("Transformation::TransformDocument(lpEcmDocument:'{0}'", lpEcmDocument.ID))
        lpErrorMessage &= Helper.FormatCallStack(ex)
        Return Nothing
      Finally
        ''LogSession.LeaveMethod(Level.Debug, Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))
      End Try

    End Function

    Public Function TransformFolder(ByVal lpFolder As Core.Folder,
                                      Optional ByRef lpErrorMessage As String = "") As Core.Folder
      If lpFolder Is Nothing Then
        Throw New ArgumentNullException("lpFolder", "lpFolder argument is null")
      End If

      Me.Target = lpFolder
      Dim lstrTransformationError As String = String.Empty
      Dim lobjActionResult As Transformations.ActionResult
      Dim lintActionCounter As Integer
      Dim lintErrorCounter As Integer
      Dim lstrLogMessage As String = Nothing

      Try

        ' In case we cancelled the previous transformation as a result of an 
        ' ExitTransformationAction we need to reset the ShouldCancel flag.  
        ' Otherwise, none of the actions will get executed.
        ShouldCancel = False

        For Each lobjAction As Action In Actions
          lstrTransformationError = String.Empty
          lintActionCounter += 1
          If Not lobjAction Is Nothing Then

            lobjAction.Transformation = Me

            ' Check to see if we have been asked to cancel out
            If ShouldCancel = True Then
              Exit For
            End If

            lobjActionResult = lobjAction.Execute(lstrTransformationError)
            If lobjActionResult.Success = False Then
              lintErrorCounter += 1
              ' Make a notation in the log
              lstrLogMessage = String.Format("Transformation action {0}. ({1}) (Document Id: {3}) failed: {2}",
                                             lintActionCounter.ToString, lobjAction.Name, lstrTransformationError, Me.Folder.Id)
              ApplicationLogging.WriteLogEntry(lstrLogMessage, TraceEventType.Warning, 32852)
              If lpErrorMessage.Length = 0 Then
                lpErrorMessage &= lstrTransformationError
              Else
                lpErrorMessage &= "; " & lstrTransformationError
              End If
            Else
              If lobjActionResult.Details.Length > 0 Then
                'If lpErrorMessage.Length = 0 Then
                '  lpErrorMessage &= lobjActionResult.Details
                'Else
                '  lpErrorMessage &= "; " & lobjActionResult.Details
                'End If
              End If
              If lstrTransformationError IsNot Nothing AndAlso lstrTransformationError.Length > 0 Then
                lstrLogMessage = String.Format("Transformation action {0}. ({1}) completed with the message: {2}",
                                             lintActionCounter.ToString, lobjAction.Name, lstrTransformationError)
                ApplicationLogging.WriteLogEntry(lstrLogMessage, TraceEventType.Warning, 32853)
              End If
            End If
          End If
        Next

        '  Delete all of the non-persistent properties created during the transformation
        'lpEcmDocument.DeleteTemporaryProperties()

        'lpEcmDocument.Versions.Sort()
        'lpEcmDocument.Versions.SortProperties()


        ' Add the transformation information to the header
        'lpEcmDocument.Header.TransformationSeries.Add(New Document.Modification(My.Computer.Name, My.User.Name, Me))

        ' Add an entry to the log
        Select Case lintErrorCounter
          Case 0
            lstrLogMessage = String.Format("Transformed folder '{0}' using transform '{1}' with {2} actions.",
                                           lpFolder.Id, Me.Name, Actions.Count)
            ApplicationLogging.WriteLogEntry(lstrLogMessage, TraceEventType.Information, 6867)
          Case 1
            lstrLogMessage = String.Format("Transformed folder '{0}' using transform '{1}' with {2} failure out of {3} actions",
                                           lpFolder.Id, Me.Name, lintErrorCounter, Actions.Count)
            ApplicationLogging.WriteLogEntry(lstrLogMessage, TraceEventType.Warning, 6868)
          Case Else
            lstrLogMessage = String.Format("Transformed folder '{0}' using transform '{1}' with {2} failures out of {3} actions",
                                           lpFolder.Id, Me.Name, lintErrorCounter, Actions.Count)
            ApplicationLogging.WriteLogEntry(lstrLogMessage, TraceEventType.Warning, 6868)
        End Select

        Return lpFolder

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("Transformation::TransformFolder(lpFolder:'{0}'", lpFolder.Id))
        lpErrorMessage &= Helper.FormatCallStack(ex)
        Return Nothing
      End Try
    End Function

    Public Function ToTransformationConfiguration() As TransformationConfiguration
      Try

        Dim lobjConfiguration As New TransformationConfiguration

        lobjConfiguration.Name = Me.Name
        lobjConfiguration.Description = Me.Description

        For Each lobjAction As Action In Me.Actions
          lobjConfiguration.Actions.Add(lobjAction.ToActionItem())
        Next

        Return lobjConfiguration

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    Friend Function DebuggerIdentifier() As String
      Try
        If Name Is Nothing OrElse Name.Length = 0 Then
          If TransformationFilePath Is Nothing OrElse TransformationFilePath.Length = 0 Then
            Return String.Format("CTS Transformation: {0} Action(s)", Actions.Count)
          Else
            Return String.Format("{0}: {1} Action(s)", TransformationFilePath, Actions.Count)
          End If
        Else
          Return String.Format("{0}: {1} Action(s)", Name, Actions.Count)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Attempts to update the transformation paths for any RunTransformation actions if necessary
    ''' </summary>
    ''' <remarks>Only usable when transformation is constructed via deserialization.</remarks>
    Private Sub ReconcileRelativeTransformationPaths()
      Try

        ' Make sure we have a value for TransformationPath, otherwise we can't use this method
        If TransformationFilePath = String.Empty Then
          ApplicationLogging.WriteLogEntry(
            "Unable to reconcile relative transformation paths, the TransformationFilePath property is not initialized.",
            TraceEventType.Error, 4404)
          Exit Sub
        End If

        For Each lobjAction As Action In Actions
          If Not lobjAction Is Nothing Then
            ' If the action is a RunTransformation action, 
            ' check to see if the transformation path was specified as a relative path.
            ' If so, we need to replace it with a fully qualified path
            If lobjAction.GetType.Name = "RunTransformationAction" Then
              Dim lobjRunTransformationAction As RunTransformationAction =
                CType(lobjAction, RunTransformationAction)

              If lobjRunTransformationAction.TransformationPath.StartsWith("\\") = False AndAlso
                lobjRunTransformationAction.TransformationPath.Substring(1, 1) <> ":" Then
                ' The transformation path was not specified as a fully qualified path
                lobjRunTransformationAction.TransformationPath =
                  RelativePath.GetRelativePath(Path.GetDirectoryName(TransformationFilePath),
                                               lobjRunTransformationAction.TransformationPath)
              End If
            End If

          End If
        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Loops through all actions and cleans up incorrectly deserialized PropertyValue nodes.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CleanUpXmlNodeArrays()
      Try

        For Each lobjAction As Action In Actions
          If lobjAction IsNot Nothing Then
            CleanUpXmlNodeArrays(lobjAction)
          End If
        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Looks for a PropertyValue node that is shown as type XmlNode() and gets the actual value instead.
    ''' </summary>
    ''' <param name="lpAction"></param>
    ''' <remarks></remarks>
    Private Sub CleanUpXmlNodeArrays(ByVal lpAction As Action)
      Try

        Dim lobjNodeArray As XmlNode() = Nothing
        Dim lobjPropertyValueInfo As Reflection.PropertyInfo
        Dim lobjPropertyValue As Object

        If lpAction Is Nothing Then
          ' Do nothing, we don't have an action to clean up
          Exit Sub
        End If

        lobjPropertyValueInfo = lpAction.GetType.GetProperty("PropertyValue")
        If lobjPropertyValueInfo IsNot Nothing Then
          lobjPropertyValue = CType(lpAction, Object).PropertyValue
          If TypeOf (lobjPropertyValue) Is XmlNode() Then
            ' We found one
            lobjNodeArray = lobjPropertyValue
            If lobjNodeArray.Length > 0 Then
              lobjPropertyValue = lobjNodeArray(0).InnerText
              CType(lpAction, Object).PropertyValue = lobjPropertyValue
              ApplicationLogging.WriteLogEntry(
                String.Format("Updated property value for action '{0}' from an XmlNode array to value '{1}'.",
                              lpAction.Name, lobjPropertyValue.ToString), TraceEventType.Information, 61431)
            End If
          End If
        End If

        ' Since decision actions have child actions we need to go through them as well
        If TypeOf (lpAction) Is DecisionAction Then
          For Each lobjAction As Action In DirectCast(lpAction, DecisionAction).TrueActions
            CleanUpXmlNodeArrays(lobjAction)
          Next
          For Each lobjAction As Action In DirectCast(lpAction, DecisionAction).FalseActions
            CleanUpXmlNodeArrays(lobjAction)
          Next
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub


#End Region

#Region "ISerialize Implementation"

    ''' <summary>
    ''' Gets the default file extension 
    ''' to be used for serialization 
    ''' and deserialization.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property DefaultFileExtension() As String Implements ISerialize.DefaultFileExtension
      Get
        Return TRANSFORMATION_FILE_EXTENSION
      End Get
    End Property

    Public Function Deserialize(ByVal lpFilePath As String, Optional ByRef lpErrorMessage As String = "") As Object Implements ISerialize.Deserialize
      Try
        Return Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        lpErrorMessage = Helper.FormatCallStack(ex)
        Return Nothing
      End Try
    End Function

    Public Function DeSerialize(ByVal lpXML As System.Xml.XmlDocument) As Object Implements ISerialize.Deserialize
      Try

        Dim lobjTransformation As Transformation = Serializer.Deserialize.XmlString(lpXML.OuterXml, Me.GetType)

        If lobjTransformation IsNot Nothing Then
          Return lobjTransformation
        Else
          Return Nothing
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Helper.DumpException(ex)
        ' Rethrow the exception
        Throw
        Return Nothing
      End Try
    End Function

    Public Function DeSerialize(ByVal lpStream As Stream) As Object
      Try
        Return Serializer.Deserialize.FromStream(lpStream, Me.GetType())
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Saves a representation of the object in an XML file.
    ''' </summary>
    Public Sub Serialize(ByVal lpFilePath As String) Implements ISerialize.Serialize
      Try
        Serialize(lpFilePath, "")
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Serialize(ByRef lpFilePath As String, ByVal lpFileExtension As String) Implements ISerialize.Serialize
      Try
        Serializer.Serialize.XmlFile(Me, lpFilePath)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Serialize(ByVal lpFilePath As String, ByVal lpWriteProcessingInstruction As Boolean, ByVal lpStyleSheetPath As String) Implements ISerialize.Serialize
      Try
        'Serializer.Serialize.XmlFile(Me, lpFilePath, , mstrXMLProcessingInstructions)
        If lpWriteProcessingInstruction = True Then
          Serializer.Serialize.XmlFile(Me, lpFilePath, , , True, lpStyleSheetPath)
        Else
          Serializer.Serialize.XmlFile(Me, lpFilePath)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function Serialize() As System.Xml.XmlDocument Implements ISerialize.Serialize
      Try
        Return Serializer.Serialize.Xml(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Public Function ToArchiveStream() As System.IO.Stream
    '  Try
    '    Return Serializer.Serialize.ToStream(Me)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    Public Function ToXmlString() As String Implements ISerialize.ToXmlString
      Try
        Return Serializer.Serialize.XmlString(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToTransformation() As Transformation Implements ITransformationConfiguration.ToTransformation
      Try
        Return Me
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToConfigurationXmlString() As String Implements ITransformationConfiguration.ToXmlString
      Try
        Return Serializer.Serialize.XmlString(Me.ToTransformationConfiguration())
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "ICloneable Implementation"

    Public Function Clone() As Object Implements System.ICloneable.Clone
      Try
        Dim lobjNewTransformation As New Transformation(Me.Serialize)
        lobjNewTransformation.Document = Nothing
        Return lobjNewTransformation
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Protected Overridable Sub InitializeLogSession() Implements ILoggable.InitializeLogSession
    '  Try
    '    mobjLogSession = ApplicationLogging.InitializeLogSession(Me.GetType.Name, System.Drawing.Color.Linen)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    'Protected Overridable Sub FinalizeLogSession() Implements ILoggable.FinalizeLogSession
    '  Try
    '    If mobjLogSession IsNot Nothing Then
    '      ApplicationLogging.FinalizeLogSession(mobjLogSession)
    '    End If
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

#End Region

#Region "ITransformation Implementation"

    <XmlIgnore()>
    Public Property ITransformationActions As ITransformationActions Implements ITransformation.Actions
      Get
        Return Me.Actions
      End Get
      Set(value As ITransformationActions)
        Me.Actions = value
      End Set
    End Property

#End Region

    '#Region "IXmlSerializable Implementation"

    '    Public Function GetSchema() As System.Xml.Schema.XmlSchema Implements System.Xml.Serialization.IXmlSerializable.GetSchema
    '      Try
    '        Return Nothing
    '      Catch ex As Exception
    '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '        ' Re-throw the exception to the caller
    '        Throw
    '      End Try
    '    End Function

    '    Public Sub ReadXml(ByVal reader As System.Xml.XmlReader) Implements System.Xml.Serialization.IXmlSerializable.ReadXml

    '      Dim lobjProcessXmlBuilder As New StringBuilder
    '      Dim lobjXmlDocument As New XmlDocument
    '      Dim lobjAttribute As XmlAttribute = Nothing

    '      Try

    '        ' <Modified by: Ernie Bahr at 11/13/2012-07:50:54 on machine: ERNIEBAHR-THINK>
    '        ' We were having problems reading when loading as part of a larger object such as JobConfiguration.  
    '        ' Recreating the xml string seems to resolve the issue.
    '        lobjProcessXmlBuilder.AppendLine(XML_HEADER)
    '        lobjProcessXmlBuilder.Append(reader.ReadOuterXml)

    '        'lobjXmlDocument.Load(reader)
    '        lobjXmlDocument.LoadXml(lobjProcessXmlBuilder.ToString)
    '        ' </Modified by: Ernie Bahr at 11/13/2012-07:50:54 on machine: ERNIEBAHR-THINK>

    '        With lobjXmlDocument
    '          ' Get the name
    '          mstrName = .DocumentElement.GetAttribute("Name")

    '          Dim lobjDescriptionNode As XmlNode = .SelectSingleNode("//Description")
    '          If lobjDescriptionNode IsNot Nothing Then
    '            mstrDescription = lobjDescriptionNode.InnerText
    '          End If

    '          ' Read the Operation elements
    '          Dim lobjActionsNode As XmlNode = .SelectSingleNode("//Actions")
    '          If lobjActionsNode IsNot Nothing Then
    '            Actions.AddRange(GetActions(lobjActionsNode))
    '          End If
    '        End With


    '        ''If reader.IsEmptyElement Then
    '        ''  reader.Read()
    '        ''  Exit Sub
    '        ''End If

    '        ''Dim lstrCurrentElementName As String = String.Empty

    '        ''' Read the Name attribute
    '        ''Me.Name = reader.GetAttribute("Name")

    '        ''Do Until reader.NodeType = XmlNodeType.EndElement AndAlso reader.Name.EndsWith("Action")
    '        ''  If reader.NodeType = XmlNodeType.Element Then
    '        ''    lstrCurrentElementName = reader.Name
    '        ''  Else
    '        ''    Select Case lstrCurrentElementName
    '        ''      ' TODO: Code for ExclusionPath and Exclusions

    '        ''      Case "Actions"
    '        ''        'Me.Actions = New TransformationActions(reader.ReadSubtree)

    '        ''      Case "Description"
    '        ''        Me.Description = reader.Value
    '        ''      Case "PropertyName"
    '        ''        Me.NewName = reader.Value
    '        ''      Case "PropertyScope"
    '        ''        Me.PropertyScope = [Enum].Parse(GetType(PropertyScope), reader.Value)
    '        ''      Case "NewName"
    '        ''        Me.NewName = reader.Value
    '        ''      Case "NewSystemName"
    '        ''        Me.NewSystemName = reader.Value
    '        ''    End Select
    '        ''  End If
    '        ''Loop
    '        ' Read the PropertyName element


    '      Catch ex As Exception
    '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '        ' Re-throw the exception to the caller
    '        Throw
    '      Finally
    '        lobjProcessXmlBuilder = Nothing
    '        lobjXmlDocument = Nothing
    '      End Try
    '    End Sub

    '    Public Shared Function GetActions(ByVal lpActionsNode As XmlNode) As Actions
    '      Try
    '        Dim lobjActions As New Actions
    '        Dim lobjAction As Action = Nothing

    '        For Each lobjActionNode As XmlNode In lpActionsNode.ChildNodes
    '          lobjAction = TransformationActionFactory.Create(lobjActionNode, CultureInfo.CurrentCulture.Name)
    '          lobjActions.Add(lobjAction)
    '        Next
    '      Catch ex As Exception

    '      End Try
    '    End Function

    '    Public Overridable Sub WriteXml(ByVal writer As System.Xml.XmlWriter) Implements System.Xml.Serialization.IXmlSerializable.WriteXml
    '      Try
    '        'WriteOperableXml(Me, writer)

    '        'With writer

    '        '  ' Write the xsi:type attribute
    '        '  .WriteAttributeString("xsi:type", "RenamePropertyAction")

    '        '  ' Write the id attribute
    '        '  .WriteAttributeString("id", Me.Id)

    '        '  ' Write the Name attribute
    '        '  .WriteAttributeString("Name", Me.Name)

    '        '  ' Write the Description attribute
    '        '  .WriteAttributeString("Description", Me.Description)

    '        '  ' Write the PropertyName element
    '        '  .WriteElementString("PropertyName", Me.PropertyName)

    '        '  ' Write the PropertyScope element
    '        '  .WriteElementString("PropertyScope", Me.PropertyScope.ToString())

    '        '  ' Write the NewName element
    '        '  .WriteElementString("NewName", Me.NewName)

    '        '  ' Optionally write the NewSystemName element
    '        '  If Not String.IsNullOrEmpty(Me.NewSystemName) Then
    '        '    .WriteElementString("NewSystemName", Me.NewSystemName)
    '        '  End If

    '        'End With

    '      Catch ex As Exception
    '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '        ' Re-throw the exception to the caller
    '        Throw
    '      End Try
    '    End Sub

    '#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
      If Not disposedValue Then
        If disposing Then
          ' TODO: dispose managed state (managed objects).
          'FinalizeLogSession()
        End If

        ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
        ' TODO: set large fields to null.
      End If
      disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
      ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
      Dispose(True)
      ' TODO: uncomment the following line if Finalize() is overridden above.
      ' GC.SuppressFinalize(Me)
    End Sub
#End Region

    'Public Property Actions1 As IActionItems Implements ITransformationConfiguration.Actions

    'Public Property Description1 As String Implements ITransformationConfiguration.Description

    'Public Property Name1 As String Implements ITransformationConfiguration.Name

    'Public Function ToXmlString1() As String Implements ITransformationConfiguration.ToXmlString

    'End Function
  End Class

End Namespace
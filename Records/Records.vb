'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports System.Runtime.Serialization
Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Records

#Region "Public Enumerations"

  ''' <summary>
  ''' Specifies the current status of a physical record
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum ChargeOutStatus

    NotChargedOut = 0
    ChargedOut = 1
    Lost = 2

  End Enum

  ''' <summary>
  ''' Records management event types
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum EventType

    External = 0
    Internal = 1
    Recurring = 2
    PredefinedDate = 3

  End Enum

  ''' <summary>
  ''' The defined types of records management entities
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum EntityType

    Invalid = -1
    Generic = 100
    RecordCategory = 101
    RecordFolder = 102
    Volume = 103
    ElectronicRecordFolder = 105
    PhysicalContainer = 106
    HybridRecordFolder = 108
    RMFolder = 109
    PhysicalRecordFolder = 110
    Record = 300
    ElectronicRecordInfo = 301
    EmailRecord = 302
    Marker = 303

  End Enum

#End Region

#Region "Public Interfaces"

  Public Interface IRecordsManager

    Enum RecordType
      ''' <summary>
      ''' Record Info
      ''' </summary>   
      Record = 300
      ''' <summary>
      ''' Electronic Record Info
      ''' </summary>
      ElectronicRecord = 301
      ''' <summary>
      ''' Email Record Info
      ''' </summary>
      EmailRecord = 302
      ''' <summary>
      ''' Marker Info
      ''' </summary>
      Marker = 303
    End Enum

    Function AddPhysicalRecord(ByRef Args As AddPhysicalRecordArgs) As Boolean
    Function DeclareRecord(ByRef Args As DeclareRecordArgs) As Boolean

  End Interface

#End Region

#Region "Public Classes"

  ''' <summary>
  ''' A base record object
  ''' </summary>
  ''' <remarks></remarks>
  Public Class RecordObject
    Inherits Core.RepositoryObject

#Region "Class Variables"

    Private menuEntityType As EntityType
    Private mstrDescription As String = ""

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' The entity type for this record
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property EntityType() As EntityType
      Get
        Return menuEntityType
      End Get
      Set(ByVal value As EntityType)
        menuEntityType = value
      End Set
    End Property

    ''' <summary>
    ''' The record description
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Description() As String
      Get
        Return mstrDescription
      End Get
      Set(ByVal value As String)
        mstrDescription = value
      End Set
    End Property

#End Region

  End Class

  ''' <summary>
  ''' The base record document class
  ''' </summary>
  ''' <remarks></remarks>
  Public Class Record
    Inherits Document

#Region "Class Variables"

    Private menuEntityType As EntityType = EntityType.Record
    Private mobjType As RecordType

#End Region

#Region "Public Properties"

    Public Property EntityType() As EntityType
      Get
        Return menuEntityType
      End Get
      Set(ByVal value As EntityType)
        menuEntityType = value
      End Set
    End Property

    ''' <summary>
    ''' The record type
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Type() As RecordType
      Get
        Return mobjType
      End Get
      Set(ByVal value As RecordType)
        mobjType = value
      End Set
    End Property

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Attempts to add the record to the File Plan
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>If the ContentSource property is not set or the 
    ''' ContentSource Provider does not implement 
    ''' IRecordsManager, an exception will be thrown.</remarks>
    Public Overloads Function Add(ByVal lpRecordFolderPath As String, ByVal lpReviewer As String) As Boolean
      Try

        Dim lstrErrorMessage As String = ""
        Dim lobjIRecordsManager As IRecordsManager = GetRMProvider()
        Dim lobjArgs As New AddPhysicalRecordArgs(Me, Me.ContentSource, lpRecordFolderPath, New ECMProperties, "CorpRecCenterDoc", lpReviewer, lstrErrorMessage)

        lobjIRecordsManager.AddPhysicalRecord(lobjArgs)
        'lobjIRecordsManager.DeclareRecord(

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Checks to see if the current content source provider implements IRecordsManager.  
    ''' If it does then it returns a reference to the provider.
    ''' </summary>
    ''' <returns>An IRecordsManager reference to the current content source provider.</returns>
    ''' <remarks>If the ContentSource property of the document is not initialized or 
    ''' the provider used by the content source does not implement the IRecordsManager 
    ''' interface, an InvalidContentSourceException will be thrown.</remarks>
    Protected Overloads Function GetRMProvider() As IRecordsManager
      Try

        ' Simply get the value from the base document implementation 
        ' and return it as a specific IRecordsManager interface reference
        '
        ' We were unable to return this interface from the base implementation
        ' Because it has no direct reference to the Records namespace
        Return CType(MyBase.GetRMProvider, IRecordsManager)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Public Overrides Function DeclareAsRecord(ByVal lpArgs As Object) As Boolean
      Try

        If lpArgs.NewRecord Is Nothing Then
          lpArgs.NewRecord = Me
        End If

        If lpArgs.RecordProperties Is Nothing Then
          lpArgs.RecordProperties = Me.Properties
        End If

        Return GetRMProvider.DeclareRecord(lpArgs)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function [Declare](ByVal lpArgs As DeclareRecordArgs) As Boolean
      Try
        Return DeclareAsRecord(lpArgs)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal lpXMLFilePath As String)
      MyBase.New(lpXMLFilePath)
      Dim lobjRecord As Record = Deserialize(lpXMLFilePath)
      Helper.AssignObjectProperties(lobjRecord, Me)
    End Sub

    Public Sub New(ByVal lpXML As Xml.XmlDocument)
      MyBase.New(lpXML)
      Dim lobjRecord As Record = DeSerialize(lpXML)
      Helper.AssignObjectProperties(lobjRecord, Me)
    End Sub

#End Region

  End Class

  ''' <summary>
  ''' A record which contains an electronic document
  ''' </summary>
  ''' <remarks></remarks>
  Public Class ElectronicRecord
    Inherits Record

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal lpXMLFilePath As String)
      MyBase.New(lpXMLFilePath)
    End Sub

    Public Sub New(ByVal lpXML As Xml.XmlDocument)
      MyBase.New(lpXML)
    End Sub

#End Region

  End Class

  ''' <summary>
  ''' A record comprised of a physical object, not an electronic document
  ''' </summary>
  ''' <remarks></remarks>
  Public Class PhysicalRecord
    Inherits Record

#Region "Class Variables"

    Private mstrBarcode As String = ""
    Private mstrMethodOfDestruction As String = ""
    Private menuChargeOutStatus As ChargeOutStatus = ChargeOutStatus.NotChargedOut
    'Private mstrChargedOutTo As String = ""
    Private mobjLocation As New Location

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' The barcode number for the physical record
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Barcode() As String
      Get
        Return mstrBarcode
      End Get
      Set(ByVal value As String)
        mstrBarcode = value
      End Set
    End Property

    ''' <summary>
    ''' Destruction instructions
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property MethodOfDestruction() As String
      Get
        Return mstrMethodOfDestruction
      End Get
      Set(ByVal value As String)
        mstrMethodOfDestruction = value
      End Set
    End Property

    ''' <summary>
    ''' The current charge out status of the physical record
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ChargeOutStatus() As ChargeOutStatus
      Get
        Return menuChargeOutStatus
      End Get
      Set(ByVal value As ChargeOutStatus)
        menuChargeOutStatus = value
      End Set
    End Property

    ''' <summary>
    ''' The current location of the record
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Location() As Location
      Get
        Return mobjLocation
      End Get
      Set(ByVal value As Location)
        mobjLocation = value
      End Set
    End Property

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Attempts to add the record to the File Plan
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>If the ContentSource property is not set or the 
    ''' ContentSource Provider does not implement 
    ''' IRecordsManager, an exception will be thrown.</remarks>
    Public Overloads Function Add(ByVal lpRecordClass As String, ByVal lpRecordFolderPath As String) As Boolean
      Try
        Return Add(lpRecordClass, lpRecordFolderPath, "Not Provided", New ECMProperties, "")
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Attempts to add the record to the File Plan
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>If the ContentSource property is not set or the 
    ''' ContentSource Provider does not implement 
    ''' IRecordsManager, an exception will be thrown.</remarks>
    Public Overloads Function Add(ByVal lpRecordClass As String, ByVal lpRecordFolderPath As String, ByVal lpReviewer As String, ByRef lpErrorMessage As String) As Boolean
      Try
        Return Add(lpRecordClass, lpRecordFolderPath, lpReviewer, New ECMProperties, lpErrorMessage)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Attempts to add the record to the File Plan
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>If the ContentSource property is not set or the 
    ''' ContentSource Provider does not implement 
    ''' IRecordsManager, an exception will be thrown.</remarks>
    Public Overloads Function Add(ByVal lpRecordClass As String,
                                ByVal lpRecordFolderPath As String,
                                ByVal lpReviewer As String,
                                ByVal lpRecordProperties As ECMProperties,
                                ByRef lpErrorMessage As String) As Boolean
      Try

        Dim lobjIRecordsManager As IRecordsManager = GetRMProvider()
        Dim lobjArgs As New AddPhysicalRecordArgs(Me, Me.ContentSource, lpRecordFolderPath, lpRecordProperties, lpRecordClass, lpReviewer, lpErrorMessage)

        Return lobjIRecordsManager.AddPhysicalRecord(lobjArgs)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal lpXMLFilePath As String)
      MyBase.New(lpXMLFilePath)
      Dim lobjPhysicalRecord As PhysicalRecord = Deserialize(lpXMLFilePath)
      Helper.AssignObjectProperties(lobjPhysicalRecord, Me)
    End Sub

    Public Sub New(ByVal lpXML As Xml.XmlDocument)
      MyBase.New(lpXML)
      Dim lobjPhysicalRecord As PhysicalRecord = Deserialize(lpXML)
      Helper.AssignObjectProperties(lobjPhysicalRecord, Me)
    End Sub

#End Region

  End Class

  Public Class DisposalTrigger
    Inherits RecordObject

#Region "Class Variables"

    Private mintCycleDays As Integer
    Private mintCycleMonths As Integer
    Private mintCycleYears As Integer
    Private mdatTriggerDate As DateTime
    Private mstrPropertyName As String
    Private mstrPropertyValue As String
    Private mobjAssociatedDisposalTrigger As DisposalTrigger
    Private mstrAggregation As String
    Private menuEventType As EventType
    Private menuOperator As Data.Criterion.pmoOperator
    Private mstrConditionXML As String
    Private mdatExternalEventOccurenceDate As DateTime

#End Region

#Region "Public Properties"

    Public Property CycleDays() As Integer
      Get
        Return mintCycleDays
      End Get
      Set(ByVal value As Integer)
        mintCycleDays = value
      End Set
    End Property

    Public Property CycleMonths() As Integer
      Get
        Return mintCycleMonths
      End Get
      Set(ByVal value As Integer)
        mintCycleMonths = value
      End Set
    End Property

    Public Property CycleYears() As Integer
      Get
        Return mintCycleYears
      End Get
      Set(ByVal value As Integer)
        mintCycleYears = value
      End Set
    End Property

    Public Property TriggerDate() As DateTime
      Get
        Return mdatTriggerDate
      End Get
      Set(ByVal value As DateTime)
        mdatTriggerDate = value
      End Set
    End Property

    Public Property PropertyName() As String
      Get
        Return mstrPropertyName
      End Get
      Set(ByVal value As String)
        mstrPropertyName = value
      End Set
    End Property

    Public Property PropertyValue() As String
      Get
        Return mstrPropertyValue
      End Get
      Set(ByVal value As String)
        mstrPropertyValue = value
      End Set
    End Property

    Public Property AssociatedDisposalTrigger() As DisposalTrigger
      Get
        Return mobjAssociatedDisposalTrigger
      End Get
      Set(ByVal value As DisposalTrigger)
        mobjAssociatedDisposalTrigger = value
      End Set
    End Property

    Public Property Aggregation() As String
      Get
        Return mstrAggregation
      End Get
      Set(ByVal value As String)
        mstrAggregation = value
      End Set
    End Property

    Public Property EventType() As EventType
      Get
        Return menuEventType
      End Get
      Set(ByVal value As EventType)
        menuEventType = value
      End Set
    End Property

    Public Property TriggerOperator() As Data.Criterion.pmoOperator
      Get
        Return menuOperator
      End Get
      Set(ByVal value As Data.Criterion.pmoOperator)
        menuOperator = value
      End Set
    End Property

    Public Property ConditionXML() As String
      Get
        Return mstrConditionXML
      End Get
      Set(ByVal value As String)
        mstrConditionXML = value
      End Set
    End Property

    Public Property ExternalEventOccurenceDate() As DateTime
      Get
        Return mdatExternalEventOccurenceDate
      End Get
      Set(ByVal value As DateTime)
        mdatExternalEventOccurenceDate = value
      End Set
    End Property

#End Region

  End Class

  Public Class DisposalSchedule
    Inherits RecordObject

#Region "Class Variables"

    'Private mReasonForChange As String
    'Private mdatCalendarDate As DateTime
    'Private mstrCutOffBase As String
    'Private mobjDisposalTrigger As DisposalTrigger
    'Private mobjDispositionCutoffAction As Object
    'Private mintDispositionEventOffsetDays As Integer
    'Private mintDispositionEventOffsetMonths As Integer
    'Private mintDispositionEventOffsetYears As Integer
    'Private mblnIsScreeningRequired As Boolean
    'Private mstrDispositionAuthority As String
    'Private mobjAssociatedRecordCategories As New List(Of String)
    'Private mobjAssociatedRecordFolders As New List(Of String)
    'Private mobjAssociatedRecordTypes As New List(Of RecordType)
    'Private mstrDispositionInstruction As String
    'Private mdatEffectiveDateModified As DateTime
    'Private mblnIsTriggerChanged As Boolean
    'Private mobjPhases As Object
    'Private menuEntityType As EntityType
    'Private mstrSweepAuditXML As String
    'Private menuSweepState As Integer

#End Region

#Region "Public Properties"

#End Region

  End Class

  Public Class RecordType
    Inherits RecordObject

#Region "Class Variables"

    'Private mobjDispositionInstructions As New List(Of DisposalSchedule)
    'Private mobjAssociatedRecord As Record
    'Private mdatDisposalScheduleAllocationDate As DateTime
    'Private mdatExternalEventOccurrenceDate As DateTime
    'Private mintRecalculatePhaseRetention As Integer
    'Private mLastModifiedDateOfSchedule As DateTime

#End Region

#Region "Public Properties"

#End Region

  End Class

  Public Class AddPhysicalRecordArgs
    Inherits DeclareRecordArgs

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal lpRecordDocument As PhysicalRecord,
                 ByVal lpFilePlanContentSource As Providers.ContentSource,
                 ByVal lpRecordFolderPath As String,
                 ByVal lpRecordProperties As ECMProperties,
                 ByVal lpRecordClass As String,
                 ByVal lpReviewer As String,
                 ByVal lpErrorMessage As String)

      MyBase.New(lpRecordDocument,
               Nothing,
               Nothing,
               lpFilePlanContentSource,
               lpRecordFolderPath,
               lpRecordProperties,
               IRecordsManager.RecordType.Marker,
               lpRecordClass,
               lpReviewer,
               lpErrorMessage,
               Nothing)

    End Sub

    Public Sub New(ByVal lpRecordDocument As PhysicalRecord,
                 ByVal lpFilePlanContentSource As Providers.ContentSource,
                 ByVal lpRecordFolderPath As String,
                 ByVal lpRecordProperties As ECMProperties,
                 ByVal lpRecordClass As String,
                 ByVal lpReviewer As String,
                 ByVal lpErrorMessage As String,
                 ByVal lpWorker As BackgroundWorker)

      MyBase.New(lpRecordDocument,
               Nothing,
               Nothing,
               lpFilePlanContentSource,
               lpRecordFolderPath,
               lpRecordProperties,
               IRecordsManager.RecordType.Marker,
               lpRecordClass,
               lpReviewer,
               lpErrorMessage,
               lpWorker)

    End Sub

#End Region

  End Class

  ''' <summary>
  ''' Contains all the parameters necessary for the DeclareRecord method.
  ''' </summary>
  ''' <remarks>Used as a sole parameter for the DeclareRecord method.</remarks>
  Public Class DeclareRecordArgs
    Inherits Arguments.BackgroundWorkerEventArgs

#Region "Class Variables"

    Private mobjSourceDocument As Document
    Private mobjContentContentSource As Providers.ContentSource
    Private mobjFilePlanContentSource As Providers.ContentSource
    Private mstrRecordFolderPath As String
    Private mobjRecordProperties As New ECMProperties
    Private menuRecordType As IRecordsManager.RecordType
    Private mstrRecordClass As String
    Private mobjRecordDocument As Record

    Private mstrReviewer As String = ""

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Reference to the newly created record object
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property NewRecord() As Record
      Get
        Return mobjRecordDocument
      End Get
      Set(ByVal value As Record)
        mobjRecordDocument = value
      End Set
    End Property

    Public Property SourceDocument() As Document
      Get
        Return mobjSourceDocument
      End Get
      Set(ByVal value As Document)
        mobjSourceDocument = value
      End Set
    End Property

    Public Property ContentContentSource() As Providers.ContentSource
      Get
        Return mobjContentContentSource
      End Get
      Set(ByVal value As Providers.ContentSource)
        mobjContentContentSource = value
      End Set
    End Property

    Public Property FilePlanContentSource() As Providers.ContentSource
      Get
        Return mobjFilePlanContentSource
      End Get
      Set(ByVal value As Providers.ContentSource)
        mobjFilePlanContentSource = value
      End Set
    End Property

    Public Property RecordFolderPath() As String
      Get
        Return mstrRecordFolderPath
      End Get
      Set(ByVal value As String)
        mstrRecordFolderPath = value
      End Set
    End Property

    Public Property RecordProperties() As ECMProperties
      Get
        Return mobjRecordProperties
      End Get
      Set(ByVal value As ECMProperties)
        mobjRecordProperties = value
      End Set
    End Property

    Public Property RecordType() As IRecordsManager.RecordType
      Get
        Return menuRecordType
      End Get
      Set(ByVal value As IRecordsManager.RecordType)
        menuRecordType = value
      End Set
    End Property

    Public Property RecordClass() As String
      Get
        Return mstrRecordClass
      End Get
      Set(ByVal value As String)
        mstrRecordClass = value
      End Set
    End Property

    Public Property Reviewer() As String
      Get
        Return mstrReviewer
      End Get
      Set(ByVal value As String)
        Try
          mstrReviewer = value
          ' Add this to the property collection
          If RecordProperties.PropertyExists("Reviewer") = False Then
            RecordProperties.Add(PropertyFactory.Create(PropertyType.ecmString, "Reviewer", mstrReviewer))
          Else
            RecordProperties("Reviewer").Value = mstrReviewer
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal lpRecordDocument As Document,
                 ByVal lpSourceDocument As Document,
                 ByVal lpContentContentSource As Providers.ContentSource,
                 ByVal lpFilePlanContentSource As Providers.ContentSource,
                 ByVal lpRecordFolderPath As String,
                 ByVal lpRecordProperties As ECMProperties,
                 ByVal lpRecordType As IRecordsManager.RecordType,
                 ByVal lpRecordClass As String,
                 ByVal lpReviewer As String,
                 ByVal lpErrorMessage As String)

      Me.New(lpRecordDocument,
           lpSourceDocument,
           lpContentContentSource,
           lpFilePlanContentSource,
           lpRecordFolderPath,
           lpRecordProperties,
           lpRecordType,
           lpRecordClass,
           lpReviewer,
           lpErrorMessage,
           Nothing)

    End Sub

    Public Sub New(ByVal lpRecordDocument As Document,
                 ByVal lpSourceDocument As Document,
                 ByVal lpContentContentSource As Providers.ContentSource,
                 ByVal lpFilePlanContentSource As Providers.ContentSource,
                 ByVal lpRecordFolderPath As String,
                 ByVal lpRecordProperties As ECMProperties,
                 ByVal lpRecordType As IRecordsManager.RecordType,
                 ByVal lpRecordClass As String,
                 ByVal lpReviewer As String,
                 ByVal lpErrorMessage As String,
                 ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpWorker)

      Try
        NewRecord = lpRecordDocument
        SourceDocument = lpSourceDocument
        ContentContentSource = lpContentContentSource
        FilePlanContentSource = lpFilePlanContentSource
        RecordFolderPath = lpRecordFolderPath
        RecordProperties = lpRecordProperties
        RecordType = lpRecordType
        RecordClass = lpRecordClass
        Reviewer = lpReviewer
        ErrorMessage = lpErrorMessage
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Sub

#End Region

  End Class

  <DataContract()>
  Public Class Location
    Inherits Core.RepositoryObject

#Region "Class Variables"

    Private mstrBarcode As String
    Private mstrReviewer As String

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' The location barcode
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember()>
    Public Property Barcode() As String
      Get
        Return mstrBarcode
      End Get
      Set(ByVal value As String)
        mstrBarcode = value
        MyBase.AddProperty(Reflection.MethodBase.GetCurrentMethod, value)
      End Set
    End Property

    ''' <summary>
    ''' The assigned reviewer
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember()>
    Public Property Reviewer() As String
      Get
        Return mstrReviewer
      End Get
      Set(ByVal value As String)
        mstrReviewer = value
        MyBase.AddProperty(Reflection.MethodBase.GetCurrentMethod, value)
      End Set
    End Property

    ''' <summary>
    ''' The location description
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DataMember()>
    Public Property Description() As String
      Get
        Return MyBase.DescriptiveText
      End Get
      Set(ByVal value As String)
        MyBase.DescriptiveText = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpName As String, ByVal lpBarcode As String, ByVal lpReviewer As String)
      Me.New(lpName, lpBarcode, lpReviewer, lpName)
    End Sub

    Public Sub New(ByVal lpName As String, ByVal lpBarcode As String, ByVal lpReviewer As String, ByVal lpDescription As String)
      MyBase.New(lpName)
      Barcode = lpBarcode
      Reviewer = lpReviewer
      Description = lpDescription
    End Sub

#End Region

  End Class

#End Region

End Namespace

'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Reflection
Imports Documents.Utilities

#End Region

Namespace Core
  ''' <summary>
  ''' Defines a document class. Returned from IClassification operations and as part of
  ''' the Repository object.
  ''' </summary>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class DocumentClass
    Inherits RepositoryObjectClass
    'Implements IStreamSerialize

#Region "Class Constants"

    ''' <summary>
    ''' Constant value representing the file extension to save XML serialized instances
    ''' to.
    ''' </summary>
    Public Const DOCUMENT_CLASS_FILE_EXTENSION As String = "dcf"

#End Region

#Region "Class Variables"

    'Private mstrID As String = String.Empty
    'Private mstrLabel As String = String.Empty
    'Private mstrName As String = String.Empty
    'Private mobjProperties As New ClassificationProperties
    'Private menuParentIdentifier As New ObjectIdentifier
    'Private mstrRepositoryName As String = String.Empty
    'Private mobjRepository As Repository
    'Private mobjSerializationSourceType As SourceType = SourceType.File

#End Region

#Region "Public Properties"

    ' ''' <summary>
    ' ''' The identifier for the document class as specified by the underlying content
    ' ''' repository.
    ' ''' </summary>
    'Public Property ID() As String
    '  Get
    '    Return mstrID
    '  End Get
    '  Set(ByVal value As String)
    '    mstrID = value
    '  End Set
    'End Property

    ' ''' <summary>The descriptive identifier for the document class.</summary>
    ' ''' <value>
    ' ''' In some cases this may be equal to the value of the ID property, otherwise it is
    ' ''' the value most suited as a descriptive value.
    ' ''' </value>
    'Public Property Label() As String
    '  Get
    '    Return mstrLabel
    '  End Get
    '  Set(ByVal value As String)
    '    mstrLabel = value
    '  End Set
    'End Property

    ' ''' <summary>The name of the document class.</summary>
    '<XmlAttribute()> _
    'Public Property Name() As String
    '  Get
    '    Return mstrName
    '  End Get
    '  Set(ByVal value As String)
    '    mstrName = value
    '  End Set
    'End Property

    ' ''' <summary>
    ' ''' Gets or sets the name of the parent repository 
    ' ''' to which this document class belongs
    ' ''' </summary>
    ' ''' <value></value>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    '<XmlAttribute("Repository")> _
    'Public Property RepositoryName() As String
    '  Get
    '    Return mstrRepositoryName
    '  End Get
    '  Set(ByVal value As String)
    '    mstrRepositoryName = value
    '  End Set
    'End Property

    '<XmlIgnore()> _
    'Public Property Repository As Repository
    '  Get
    '    Return mobjRepository
    '  End Get
    '  Set(ByVal value As Repository)
    '    Try
    '      mobjRepository = value
    '      If mobjRepository IsNot Nothing AndAlso mobjRepository.Name IsNot Nothing Then
    '        mstrRepositoryName = mobjRepository.Name
    '      End If
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Set
    'End Property

    ' ''' <summary>The document properties associated with the document class.</summary>
    'Public Property Properties() As ClassificationProperties
    '  Get
    '    Return mobjProperties
    '  End Get
    '  Set(ByVal value As ClassificationProperties)
    '    mobjProperties = value
    '  End Set
    'End Property

    ' ''' <summary>
    ' ''' If the DocumentClass is a child class then this value would refer to the parent class to which this DocumentClass is inherited.
    ' ''' </summary>
    ' ''' <value></value>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Property ParentIdentifier() As ObjectIdentifier
    '  Get
    '    Return menuParentIdentifier
    '  End Get
    '  Set(ByVal value As ObjectIdentifier)
    '    menuParentIdentifier = value
    '  End Set
    'End Property

#End Region

#Region "Constructors"

    ''' <summary>Default Constructor</summary>
    Public Sub New()
      MyBase.New()
    End Sub

    ''' <summary>
    ''' Constructs a new document class object using the specified name, properties
    ''' collection, ID and label.
    ''' </summary>
    Public Sub New(ByVal lpName As String,
                   ByVal lpProperties As ClassificationProperties,
                   Optional ByVal lpId As String = "",
                   Optional ByVal lpLabel As String = "")
      MyBase.New(lpName, lpProperties, lpId, lpLabel)
      'Try
      '  Name = lpName
      '  Properties = lpProperties
      '  ID = lpId
      '  Label = lpLabel
      'Catch ex As Exception
      '  ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      'End Try

    End Sub

    ''' <summary>
    ''' Constructs a new document class object using the specified stream.
    ''' </summary>
    ''' <param name="lpStream">An IO.Stream object derived from a document class file.</param>
    ''' <remarks>
    ''' If there are child choice lists, they can't be retrieved 
    ''' using this constructor.  Use New(lpStream, lpZipFile) instead.
    '''  The choice lists are actually stored in separate files that 
    ''' can only be from the zipped repository archive file.
    '''</remarks>
    Public Sub New(ByVal lpStream As IO.Stream)
      MyBase.New(lpStream)
      'Try
      '  Dim lobjDocumentClass As DocumentClass = Deserialize(lpStream)
      '  Helper.AssignObjectProperties(lobjDocumentClass, Me)
      'Catch ex As Exception
      '  ApplicationLogging.LogException(ex, String.Format("DocumentClass::New('{0}')", lpStream))
      '  ' Re-throw the exception to the caller
      '  Throw
      'End Try
    End Sub

#End Region

#Region "Public Methods"

    <Obsolete("This method has been deprecated, use FindInvalidProperties instead.")>
    Public Function FindInvalidClassProperties(ByVal lpDocument As Document,
                                           ByVal lpPropertyScope As PropertyScope,
                                           Optional ByRef lpErrorMessage As String = "") As IProperties
      Try
        Return FindInvalidProperties(lpDocument, lpPropertyScope, lpErrorMessage)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Evaluate the properties of the document to see which ones 
    ''' are invalid for the current document class.
    ''' </summary>
    ''' <param name="lpDocument"></param>
    ''' <returns>Properties collection of invalid properties for current class</returns>
    ''' <remarks></remarks>
    Public Function FindInvalidProperties(ByVal lpDocument As Document,
                                               ByVal lpPropertyScope As PropertyScope,
                                               Optional ByRef lpErrorMessage As String = "") As IProperties
      Try
        Return FindInvalidProperties(lpDocument, lpPropertyScope, True, lpErrorMessage)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Evaluate the properties of the document to see which ones 
    ''' are invalid for the current document class.
    ''' </summary>
    ''' <param name="lpDocument"></param>
    ''' <returns>Properties collection of invalid properties for current class</returns>
    ''' <remarks></remarks>
    Public Function FindInvalidProperties(ByVal lpDocument As Document,
                                               ByVal lpPropertyScope As PropertyScope,
                                               ByVal lpWarnForNoncompliance As Boolean,
                                               Optional ByRef lpErrorMessage As String = "") As IProperties

      Try
        ' Save the names of the properties we delete
        Dim lstrDeletedProperties As String = "Deleted Properties: "
        Dim lstrDeletedPropertyDelimiter As String = ","
        Dim lobjFoundProperty As ClassificationProperty = Nothing
        Dim lobjInvalidProperties As New InvalidProperties
        Dim lintVersionCounter As Integer = 0

        ' Make sure the DocumentClass exists in the collection
        Dim lstrDocumentClassName As String
        lstrDocumentClassName = lpDocument.DocumentClass '.Properties("Document Class").Value
        If lstrDocumentClassName.Length = 0 Then
          lpErrorMessage = "DocumentClass not specified"
          ApplicationLogging.WriteLogEntry(String.Format("DocumentClass::FindInvalidProperties:{0}", lpErrorMessage), TraceEventType.Warning)
          Return lobjInvalidProperties
        End If

        ' Set the properties to exclude
        Dim lstrExcludeProperties(7) As String
        lstrExcludeProperties(0) = "Document Class"
        lstrExcludeProperties(1) = "DocumentClass"
        lstrExcludeProperties(2) = "Folder Path"
        lstrExcludeProperties(3) = "FolderPath"
        lstrExcludeProperties(4) = "Folders Filed In"
        lstrExcludeProperties(5) = "FoldersFiledIn"
        lstrExcludeProperties(6) = "ObjectID"
        Dim lblnIsExclusionProperty As Boolean = False

        Select Case lpPropertyScope
          Case PropertyScope.DocumentProperty
            ' Only clear the document level properties
            For Each lobjProperty As ECMProperty In lpDocument.Properties
              If Properties.PropertyExists(lobjProperty.Name, lobjFoundProperty) = False Then
                lblnIsExclusionProperty = False
                For lintExclusionCounter As Integer = 0 To lstrExcludeProperties.Length - 1
                  If String.Compare(lobjProperty.Name, lstrExcludeProperties(lintExclusionCounter),
                                    StringComparison.InvariantCultureIgnoreCase) = 0 Then
                    lblnIsExclusionProperty = True
                    Exit For
                  End If
                Next
                If lblnIsExclusionProperty = False Then
                  If (Not lobjInvalidProperties.PropertyExists(lobjProperty.Name)) Then
                    lstrDeletedProperties &= lobjProperty.Name & lstrDeletedPropertyDelimiter
                    lobjInvalidProperties.Add(lobjProperty, InvalidProperty.InvalidPropertyScope.Document)
                  End If
                End If
              Else
                ' See if the property is settable
                If lobjFoundProperty.Settability = ClassificationProperty.SettabilityEnum.READ_ONLY OrElse
                lobjFoundProperty.Settability = ClassificationProperty.SettabilityEnum.SETTABLE_ONLY_ON_CREATE Then
                  If (Not lobjInvalidProperties.PropertyExists(lobjProperty.Name)) Then
                    lstrDeletedProperties &= lobjProperty.Name & lstrDeletedPropertyDelimiter
                    lobjInvalidProperties.Add(lobjProperty, InvalidProperty.InvalidPropertyScope.Document)
                  End If
                End If
              End If
            Next
          Case PropertyScope.VersionProperty
            ' Only clear the version level properties
            For Each lobjversion As Version In lpDocument.Versions
              For Each lobjProperty As ECMProperty In lobjversion.Properties
                If Properties.PropertyExists(lobjProperty.Name, lobjFoundProperty) = False Then
                  If lobjInvalidProperties.PropertyExists(lobjProperty.Name) = False Then
                    ' <Modified by: Ernie at 1/3/2014-2:00:54 PM on machine: ERNIE-THINK>
                    ' <This was filtered to suppress the log warnings for MajorVersion for FogBugz case 341 from Tilo Sönnichsen for TransCanada.>
                    If lobjProperty.Name <> "MajorVersion" Then
                      lstrDeletedProperties &= lobjProperty.Name & lstrDeletedPropertyDelimiter
                    End If
                    ' </This was filtered to suppress the log warnings for MajorVersion for FogBugz case 341 from Tilo Sönnichsen for TransCanada.>
                    ' </Modified by: Ernie at 1/3/2014-2:00:54 PM on machine: ERNIE-THINK>
                    lobjInvalidProperties.Add(lobjProperty, InvalidProperty.InvalidPropertyScope.AllVersions)
                  End If
                Else
                  Select Case lobjFoundProperty.Settability
                    Case ClassificationProperty.SettabilityEnum.READ_ONLY
                      If (Not lobjInvalidProperties.PropertyExists(lobjProperty.Name)) Then
                        lstrDeletedProperties &= lobjProperty.Name & lstrDeletedPropertyDelimiter
                        lobjInvalidProperties.Add(lobjProperty, InvalidProperty.InvalidPropertyScope.AllVersions)
                      End If
                      'Case ClassificationProperty.SettabilityEnum.SETTABLE_ONLY_ON_CREATE
                      '  If (Not lobjInvalidProperties.PropertyExists(lobjProperty.Name)) Then
                      '    ' <Removed by: Ernie at: 1/3/2014-1:54:57 PM on machine: ERNIE-THINK>
                      '    ' <This was commented out to suppress the log warnings for Creator and DateCreated for FogBugz case 341 from Tilo Sönnichsen for TransCanada.>
                      '    ' lstrDeletedProperties &= lobjProperty.Name & lstrDeletedPropertyDelimiter
                      '    ' </This was commented out to suppress the log warnings for Creator and DateCreated for FogBugz case 341 from Tilo Sönnichsen for TransCanada.>
                      '    ' </Removed by: Ernie at: 1/3/2014-1:54:57 PM on machine: ERNIE-THINK>                        lobjInvalidProperties.Add(lobjProperty, InvalidProperty.InvalidPropertyScope.AllExceptFirstVersion)
                      '  End If
                  End Select
                End If
              Next
              lintVersionCounter += 1
            Next
          Case PropertyScope.BothDocumentAndVersionProperties
            ' Clear both the document and version level properties
            ' First the document level properties
            For Each lobjProperty As ECMProperty In lpDocument.Properties
              If Properties.PropertyExists(lobjProperty, lobjFoundProperty) = False Then
                lblnIsExclusionProperty = False
                For lintExclusionCounter As Integer = 0 To lstrExcludeProperties.Length - 1
                  If lobjProperty.Name = lstrExcludeProperties(lintExclusionCounter) Then
                    lblnIsExclusionProperty = True
                    Exit For
                  End If
                Next
                If lblnIsExclusionProperty = False Then
                  If lobjInvalidProperties.PropertyExists(lobjProperty.Name) = False Then
                    lstrDeletedProperties &= lobjProperty.Name & lstrDeletedPropertyDelimiter
                    lobjInvalidProperties.Add(lobjProperty, InvalidProperty.InvalidPropertyScope.Document)
                  End If
                End If
              Else
                ' See if the property is settable
                Select Case lobjFoundProperty.Settability
                  Case ClassificationProperty.SettabilityEnum.READ_ONLY
                    If lobjInvalidProperties.PropertyExists(lobjProperty.Name) = False Then
                      lstrDeletedProperties &= lobjProperty.Name & lstrDeletedPropertyDelimiter
                      lobjInvalidProperties.Add(lobjProperty, InvalidProperty.InvalidPropertyScope.AllVersions)
                    End If
                    'Case ClassificationProperty.SettabilityEnum.SETTABLE_ONLY_ON_CREATE
                    '  If lobjInvalidProperties.PropertyExists(lobjProperty.Name) = False Then
                    '    ' <Removed by: Ernie at: 1/3/2014-1:54:57 PM on machine: ERNIE-THINK>
                    '    ' <This was commented out to suppress the log warnings for Creator and DateCreated for FogBugz case 341 from Tilo Sönnichsen for TransCanada.>
                    '    ' lstrDeletedProperties &= lobjProperty.Name & lstrDeletedPropertyDelimiter
                    '    ' </This was commented out to suppress the log warnings for Creator and DateCreated for FogBugz case 341 from Tilo Sönnichsen for TransCanada.>
                    '    ' </Removed by: Ernie at: 1/3/2014-1:54:57 PM on machine: ERNIE-THINK>
                    '    lobjInvalidProperties.Add(lobjProperty, InvalidProperty.InvalidPropertyScope.AllExceptFirstVersion)
                    '  End If
                End Select
              End If
            Next
            ' Now the version level properties
            For Each lobjversion As Version In lpDocument.Versions
              For Each lobjProperty As ECMProperty In lobjversion.Properties
                If Properties.PropertyExists(lobjProperty, lobjFoundProperty) = False Then
                  If lobjInvalidProperties.PropertyExists(lobjProperty.Name) = False Then
                    ' <Modified by: Ernie at 1/3/2014-2:00:54 PM on machine: ERNIE-THINK>
                    ' <This was filtered to suppress the log warnings for MajorVersion for FogBugz case 341 from Tilo Sönnichsen for TransCanada.>
                    If lobjProperty.Name <> "MajorVersion" Then
                      lstrDeletedProperties &= lobjProperty.Name & lstrDeletedPropertyDelimiter
                    End If
                    ' </This was filtered to suppress the log warnings for MajorVersion for FogBugz case 341 from Tilo Sönnichsen for TransCanada.>
                    ' </Modified by: Ernie at 1/3/2014-2:00:54 PM on machine: ERNIE-THINK>
                    lobjInvalidProperties.Add(lobjProperty, InvalidProperty.InvalidPropertyScope.AllVersions)
                  End If
                Else
                  Select Case lobjFoundProperty.Settability
                    Case ClassificationProperty.SettabilityEnum.READ_ONLY
                      If lobjInvalidProperties.PropertyExists(lobjProperty.Name) = False Then
                        lstrDeletedProperties &= lobjProperty.Name & lstrDeletedPropertyDelimiter
                        lobjInvalidProperties.Add(lobjProperty, InvalidProperty.InvalidPropertyScope.AllVersions)
                      End If
                      'Case ClassificationProperty.SettabilityEnum.SETTABLE_ONLY_ON_CREATE
                      '  If lobjInvalidProperties.PropertyExists(lobjProperty.Name) = False Then
                      '    ' <Removed by: Ernie at: 1/3/2014-1:54:57 PM on machine: ERNIE-THINK>
                      '    ' <This was commented out to suppress the log warnings for Creator and DateCreated for FogBugz case 341 from Tilo Sönnichsen for TransCanada.>
                      '    ' lstrDeletedProperties &= lobjProperty.Name & lstrDeletedPropertyDelimiter
                      '    ' </This was commented out to suppress the log warnings for Creator and DateCreated for FogBugz case 341 from Tilo Sönnichsen for TransCanada.>
                      '    ' </Removed by: Ernie at: 1/3/2014-1:54:57 PM on machine: ERNIE-THINK>
                      '    lobjInvalidProperties.Add(lobjProperty, InvalidProperty.InvalidPropertyScope.AllExceptFirstVersion)
                      '  End If
                  End Select
                End If
              Next
              lintVersionCounter += 1
            Next
        End Select

        lstrDeletedProperties = lstrDeletedProperties.Replace("Deleted Properties: ", String.Empty)

        If Not String.IsNullOrEmpty(lstrDeletedProperties) Then

          ApplicationLogging.WriteLogEntry(String.Format("DocumentClass::FindInvalidProperties - DeletedProperties:{0}",
                                           lstrDeletedProperties), TraceEventType.Verbose, 61200)

          If lpWarnForNoncompliance Then

            ApplicationLogging.WriteLogEntry(String.Format("The properties ({0}) are invalid for target document class '{1}'.",
                                             lstrDeletedProperties, Me.Name), Reflection.MethodBase.GetCurrentMethod,
                                             TraceEventType.Warning, 61203)
          End If

        End If

#If SILVERLIGHT <> 1 Then
        ' Sort the results by Name
        lobjInvalidProperties.Sort()
#End If

        Return lobjInvalidProperties

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("DocumentClass::FindInvalidProperties('{0}', '{1}', '{2}')",
                                                          lpDocument.ID, lpPropertyScope.ToString, lpErrorMessage))
        Return Nothing
      End Try

    End Function

    ''' <summary>
    ''' Evaluate the properties of the source document to see which ones 
    ''' are invalid for the destination document class.
    ''' </summary>
    ''' <returns>Properties collection of invalid properties for current class</returns>
    ''' <remarks></remarks>
    Public Shared Function FindInvalidProperties(ByVal lpSourceDocumentClass As DocumentClass,
                                                  ByVal lpDestinationDocumentClass As DocumentClass,
                                                  Optional ByRef lpErrorMessage As String = "") As IProperties

      Try
        ' Save the names of the properties we delete
        ' Dim lstrInvalidProperties As String = "Invalid Properties: "
        Dim lstrDeletedPropertyDelimiter As String = ","
        Dim lobjFoundProperty As ClassificationProperty = Nothing
        Dim lobjInvalidProperties As New InvalidProperties
        Dim lintVersionCounter As Integer = 0


        ' Set the properties to exclude
        Dim lstrExcludeProperties(7) As String
        lstrExcludeProperties(0) = "Document Class"
        lstrExcludeProperties(1) = "DocumentClass"
        lstrExcludeProperties(2) = "Folder Path"
        lstrExcludeProperties(3) = "FolderPath"
        lstrExcludeProperties(4) = "Folders Filed In"
        lstrExcludeProperties(5) = "FoldersFiledIn"
        lstrExcludeProperties(6) = "ObjectID"
        Dim lblnIsExclusionProperty As Boolean = False

        For Each lobjProperty As ClassificationProperty In lpSourceDocumentClass.Properties
          If lpDestinationDocumentClass.Properties.PropertyExists(lobjProperty.Name, lobjFoundProperty) Then
            ' See if the property is settable
            If lobjFoundProperty.Settability = ClassificationProperty.SettabilityEnum.READ_ONLY OrElse
            lobjFoundProperty.Settability = ClassificationProperty.SettabilityEnum.SETTABLE_ONLY_ON_CREATE Then
              If (Not lobjInvalidProperties.PropertyExists(lobjProperty.Name)) Then
                ' lstrInvalidProperties &= lobjProperty.Name & lstrDeletedPropertyDelimiter
                lobjInvalidProperties.Add(lobjProperty, InvalidProperty.InvalidPropertyScope.Document)
              End If
            End If
            ' See if the cardinality is the same
            If lobjFoundProperty.Cardinality <> lobjProperty.Cardinality Then
              If (Not lobjInvalidProperties.PropertyExists(lobjProperty.Name)) Then
                ' lstrInvalidProperties &= lobjProperty.Name & lstrDeletedPropertyDelimiter
                lobjInvalidProperties.Add(lobjProperty, InvalidProperty.InvalidPropertyScope.Document)
              End If
            End If
          Else
            lblnIsExclusionProperty = False
            For lintExclusionCounter As Integer = 0 To lstrExcludeProperties.Length - 1
              If String.Compare(lobjProperty.Name, lstrExcludeProperties(lintExclusionCounter),
                                StringComparison.InvariantCultureIgnoreCase) = 0 Then
                lblnIsExclusionProperty = True
                Exit For
              End If
            Next
            If lblnIsExclusionProperty = False Then
              If (Not lobjInvalidProperties.PropertyExists(lobjProperty.Name)) Then
                ' lstrInvalidProperties &= lobjProperty.Name & lstrDeletedPropertyDelimiter
                lobjInvalidProperties.Add(lobjProperty, InvalidProperty.InvalidPropertyScope.Document)
              End If
            End If
          End If
        Next

        'If lstrInvalidProperties.EndsWith(lstrDeletedPropertyDelimiter) Then
        '  lstrInvalidProperties = lstrInvalidProperties.Remove(lstrInvalidProperties.Length - lstrDeletedPropertyDelimiter.Length)
        'End If

#If SILVERLIGHT <> 1 Then
        ' Sort the results by Name
        lobjInvalidProperties.Sort()
#End If

        Return lobjInvalidProperties

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Public Function IsSubscribed(lpProperty As IProperty)
      Try
        Return Properties.Contains(lpProperty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Public Overrides Function ToString() As String
    '  Try
    '    Return Me.Name
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    Return Me.GetType.FullName
    '  End Try
    'End Function

#End Region

    '#Region "Private Methods"

    '    Protected Overridable Function DebuggerIdentifier() As String
    '      Try
    '        Dim lobjIdentifierBuilder As New StringBuilder

    '        If Not String.IsNullOrEmpty(ID) Then
    '          lobjIdentifierBuilder.AppendFormat("{0}: ", ID)
    '        End If

    '        If Not String.IsNullOrEmpty(Label) Then
    '          lobjIdentifierBuilder.AppendFormat("{0}", Label)
    '        ElseIf Not String.IsNullOrEmpty(Name) Then
    '          lobjIdentifierBuilder.AppendFormat("{0}", Name)
    '        End If

    '        If Properties IsNot Nothing Then
    '          If Properties.Count > 0 Then
    '            lobjIdentifierBuilder.AppendFormat(" ({0} Properties)", Properties.Count)
    '          Else
    '            lobjIdentifierBuilder.Append(" (No Properties)")
    '          End If
    '        Else
    '          lobjIdentifierBuilder.Append(" (No Properties)")
    '        End If

    '        Return lobjIdentifierBuilder.ToString

    '      Catch ex As Exception
    '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '        ' Re-throw the exception to the caller
    '        Throw
    '      End Try
    '    End Function

    '#End Region

    '#Region "IStreamSerialize Implementation"

    '    Public Function DeSerialize(ByVal lpStream As System.IO.Stream) As Object Implements IStreamSerialize.DeSerialize
    '      Try
    '        Return Serializer.Deserialize.FromStream(lpStream, Me.GetType)
    '      Catch ex As Exception
    '        ApplicationLogging.LogException(ex, String.Format("{0}::Deserialize(lpXML)", Me.GetType.Name))
    '        Helper.DumpException(ex)
    '        Return Nothing
    '      End Try
    '    End Function

    '    Public Function SerializeToStream() As System.IO.Stream Implements IStreamSerialize.SerializeToStream
    '      Try
    '        Return Serializer.Serialize.ToStream(Me)
    '      Catch ex As Exception
    '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '        '  Re-throw the exception to the caller
    '        Throw
    '      End Try
    '    End Function

    '#End Region

  End Class

End Namespace
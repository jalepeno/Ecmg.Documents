'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  Folder.vb
'   Description :  [type_description_here]
'   Created     :  3/5/2015 1:59:10 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Collections.Specialized
Imports System.Xml.Serialization
Imports Documents.Configuration
Imports Documents.Exceptions
Imports Documents.Providers
Imports Documents.SerializationUtilities
Imports Documents.Transformations
Imports Documents.Utilities
Imports Newtonsoft.Json

#End Region

Namespace Core

  <Serializable()>
  Public Class Folder
    Inherits RepositoryObject
    Implements IFolder
    Implements IDisposable
    'Implements ILoggable

#Region "Public Events"

    Public Event FolderSelected(sender As Object, e As Arguments.FolderEventArgs) Implements IFolder.FolderSelected

#End Region

#Region "Class Constants"

    ''' <summary>
    ''' Constant value representing the file extension to save XML serialized instances
    ''' to.
    ''' </summary>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public Const CONTENT_FOLDER_FILE_EXTENSION As String = "cff"

#End Region

#Region "Class Variables"

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private mobjAuditEvents As New AuditEvents
    <NonSerializedAttribute()>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private mobjContentSource As ContentSource
    Private mstrFolders As Collections.Specialized.StringCollection
    Private mstrPath As String = String.Empty
    Private mstrLabel As String = String.Empty
    <NonSerializedAttribute()>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private mobjProvider As IProvider

#End Region

#Region "Public Properties"

    <JsonIgnore()>
    Public ReadOnly Property IsDisposed() As Boolean
      Get
        Return disposedValue
      End Get
    End Property

    <JsonIgnore()>
    Public ReadOnly Property DefaultFileExtension() As String
      Get
        Return CONTENT_FOLDER_FILE_EXTENSION
      End Get
    End Property

    <XmlAttribute()>
    Public Property FolderClass As String
      Get
        Try

          If IsDisposed Then
            Throw New ObjectDisposedException(Me.GetType.ToString)
          End If

          If Properties.PropertyExists("Folder Class") = False Then
            Properties.Add(PropertyFactory.Create(PropertyType.ecmString,
                                                  "Folder Class", String.Empty))
          End If
          Return Properties("Folder Class").Value.ToString
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Return String.Empty
        End Try
      End Get

      Set(ByVal Value As String)
        Try

          If IsDisposed Then
            Throw New ObjectDisposedException(Me.GetType.ToString)
          End If

          If Properties.PropertyExists("Folder Class") Then
            Properties("Folder Class").Value = Value
          Else
            Properties.Add(PropertyFactory.Create(PropertyType.ecmString,
                                      "Folder Class", Value))

          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("Document::Set_DocumentClass('{0}')", Value))
        End Try
      End Set
    End Property

    Public Property AuditEvents As Core.AuditEvents Implements Core.IAuditableItem.AuditEvents
      Get
        Try
          Return mobjAuditEvents
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Core.AuditEvents)
        Try
          mobjAuditEvents = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property Identifier As String Implements Core.IMetaHolder.Identifier
      Get
        Try
          Return Me.Id
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          Me.Id = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    <XmlIgnore()>
    <JsonIgnore()>
    Public Overloads Property Metadata As IProperties Implements IFolder.Metadata
      Get
        Try
          Return Me.Properties
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As IProperties)
        Try
          Me.Properties = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public ReadOnly Property ContentCount As Long Implements IFolder.ContentCount
      Get
        Try
          ' This class is focused on the folder object itself, not the content.
          Return 0
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property ContentCount(FolderRecursionLevel As Core.RecursionLevel) As Long Implements IFolder.ContentCount
      Get
        Try
          ' This class is focused on the folder object itself, not the content.
          Return 0
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    <XmlIgnore()>
    <JsonIgnore()>
    Public ReadOnly Property Contents As Core.FolderContents Implements IFolder.Contents
      Get
        Try
          ' This class is focused on the folder object itself, not the content.
          Return New FolderContents()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ''' <summary>
    ''' Used to perform repository operations on the folder.  
    ''' </summary>
    ''' <value></value>
    ''' <returns>A ContentSource object reference</returns>
    ''' <remarks>If set to a valid content source object in which the folder resides, 
    ''' and the provider associated with the content source implements the "IBasicContentServicesProvider" interface, 
    ''' the document can perform the operations exposed by that interface.</remarks>
    <XmlIgnore()>
    <JsonIgnore()>
    Public Property ContentSource As ContentSource Implements IFolder.ContentSource
      Get
        Try
          If IsDisposed Then
            Throw New ObjectDisposedException(Me.GetType.ToString)
          End If

          Return mobjContentSource
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As ContentSource)
        Try
          If IsDisposed Then
            Throw New ObjectDisposedException(Me.GetType.ToString)
          End If

          mobjContentSource = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public ReadOnly Property FolderNames As Specialized.StringCollection Implements IFolder.FolderNames
      Get
        Try
          Return mstrFolders
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Property InvisiblePassThrough As Boolean Implements IFolder.InvisiblePassThrough
      Get
        Try
          Return False
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Boolean)
        Try
          ' Ignore
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property Label As String Implements IFolder.Label
      Get
        Try
          Return mstrLabel
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As String)
        Try
          mstrLabel = value
          AddProperty(Reflection.MethodBase.GetCurrentMethod, value)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property MaxContentCount As Long Implements IFolder.MaxContentCount
      Get
        Return 0
      End Get
      Set(value As Long)
        ' Ignore
      End Set
    End Property

    Public Overloads Property Name As String Implements IFolder.Name
      Get
        Try
          Return MyBase.Name
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          MyBase.Name = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public ReadOnly Property Path As String Implements IFolder.Path
      Get
        Try
          Return mstrPath
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    <XmlIgnore(), JsonIgnore()>
    Public Property Provider As IProvider Implements IFolder.Provider
      Get
        Try
          Return mobjProvider
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As IProvider)
        Try
          mobjProvider = value
          AddProperty(Reflection.MethodBase.GetCurrentMethod, value, PropertyType.ecmObject)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public ReadOnly Property SubFolderCount As Long Implements IFolder.SubFolderCount
      Get
        Try
          Return 0
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    <XmlIgnore()>
    <JsonIgnore()>
    Public ReadOnly Property SubFolders As Core.Folders Implements IFolder.SubFolders
      Get
        Try
          Return New Folders
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property TreeLabel As String Implements IFolder.TreeLabel
      Get
        Try
          If String.IsNullOrEmpty(mstrLabel) Then
            Return Name
          Else
            Return String.Format("{0} ({1})", Label, Name)
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      Try
        Initialize()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpFilePath As String)
      Try
        Dim lobjFolder As Folder = Serializer.Deserialize.XmlFile(lpFilePath, GetType(Folder))
        Helper.AssignObjectProperties(lobjFolder, Me)
        mstrPath = lobjFolder.Path
        Initialize()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Function GetContentCount() As Long Implements IFolder.GetContentCount
      Try
        Return 0
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetContentCount(lpFolderRecursionLevel As Core.RecursionLevel) As Long Implements IFolder.GetContentCount
      Try
        Return 0
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetFolderByPath(lpFolderPath As String, lpMaxContentCount As Long) As IFolder Implements IFolder.GetFolderByPath
      Try
        Throw New NotImplementedException
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetID() As String Implements IFolder.GetID
      Try
        Return Me.Id
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetPath() As String Implements IFolder.GetPath
      Try
        Return Me.Path
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Friend Sub ChangePath(lpNewPath As String)
      Try
        mstrPath = lpNewPath

        If Properties.PropertyExists("Path") Then
          ChangePropertyValue("Path", lpNewPath)
        End If

        If Properties.PropertyExists("Path Name") Then
          ChangePropertyValue("Path Name", lpNewPath)
        End If

        Dim lobjFolderTree As New FolderTree(lpNewPath)

        Name = lobjFolderTree.Name

        If Properties.PropertyExists("Name") Then
          ChangePropertyValue("Name", lobjFolderTree.Name)
        End If
        If Properties.PropertyExists("Folder Name") Then
          ChangePropertyValue("Folder Name", lobjFolderTree.Name)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    '''     Modifies the folder path by replacing the 'old' path with the 'new' path.  
    '''   This can be used to replace all or any portion of the folder path.
    ''' </summary>
    ''' <param name="lpOldPath" type="String">
    '''     <para>
    '''         The portion of the path to replace.
    '''     </para>
    ''' </param>
    ''' <param name="lpNewPath" type="String">
    '''     <para>
    '''         The value to substitute.
    '''     </para>
    ''' </param>
    Friend Sub ReplacePath(lpOldPath As String, lpNewPath As String)
      Try
        'LogSession.EnterMethod(Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))
        Dim lstrOriginalPath As String = GetPathValue()
        ApplicationLogging.WriteLogEntry(String.Format("OriginalPath: '{0}'", lstrOriginalPath),
                         Reflection.MethodBase.GetCurrentMethod(), TraceEventType.Information, 24485)

        'LogSession.LogCollection(Level.Message, "OriginalPath ASCII Characters", ECMProperty.StringToASCIIKeyMap(lstrOriginalPath))

        If lstrOriginalPath.Contains(lpOldPath) Then
          Dim lstrNewPath As String = lstrOriginalPath.Replace(lpOldPath, lpNewPath)

          ApplicationLogging.WriteLogEntry(String.Format("ReplacePath('{0}','{1}' -> '{2}')", lpOldPath, lpNewPath, lstrNewPath),
                           Reflection.MethodBase.GetCurrentMethod(), TraceEventType.Information, 24486)

          'LogSession.LogCollection(Level.Message, "OldPath ASCII Characters", ECMProperty.StringToASCIIKeyMap(lpOldPath))
          'LogSession.LogCollection(Level.Message, "NewPath ASCII Characters", ECMProperty.StringToASCIIKeyMap(lpNewPath))
          'LogSession.LogCollection(Level.Message, "CombinedPath ASCII Characters", ECMProperty.StringToASCIIKeyMap(lstrNewPath))

          ApplicationLogging.WriteLogEntry(String.Format("Replacing '{0}' with '{1}'.", lstrOriginalPath, lstrNewPath),
                                   Reflection.MethodBase.GetCurrentMethod(), TraceEventType.Information, 24487)

          ChangePath(lstrNewPath)

          mstrPath = lstrNewPath
        Else
          'LogSession.LogMessage(String.Format("The original path '{0}' does not contain the value to replace '{1}'.", lstrOriginalPath, lpOldPath))
          'LogSession.LogCollection(Level.Message, "OldPath ASCII Characters", ECMProperty.StringToASCIIKeyMap(lpOldPath))
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      Finally
        'LogSession.LeaveMethod(Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))
      End Try
    End Sub

    Friend Sub RemoveFolderLevel(lpTargetFolderLevel As Integer)
      Try
        Dim lstrOriginalPath As String = GetPathValue()
        ApplicationLogging.WriteLogEntry(String.Format("OriginalPath: '{0}'", lstrOriginalPath),
                         Reflection.MethodBase.GetCurrentMethod(), TraceEventType.Information, 24585)

        ''LogSession.LogCollection(Level.Message, "OriginalPath ASCII Characters", ECMProperty.StringToASCIIKeyMap(lstrOriginalPath))
        Dim lobjFolderTree As New FolderTree(lstrOriginalPath)
        Dim lobjShortenedFolderTree As FolderTree = lobjFolderTree.RemoveLevel(lpTargetFolderLevel)

        Dim lstrNewPath As String = lobjShortenedFolderTree.FolderPath

        ApplicationLogging.WriteLogEntry(String.Format("Replacing '{0}' with '{1}'.", lstrOriginalPath, lstrNewPath),
                                 Reflection.MethodBase.GetCurrentMethod(), TraceEventType.Information, 24587)

        ChangePath(lstrNewPath)

        mstrPath = lstrNewPath

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Sub

    Public Function GetProperty(ByVal lpPropertyName As String) As ECMProperty
      Try
        Return Properties(lpPropertyName)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        Throw New InvalidOperationException(String.Format("Unable to Get Property '{0}'", lpPropertyName), ex)
      End Try
    End Function

    ''' <summary>
    ''' Changes the value of the specifed property in the folder
    ''' </summary>
    ''' <param name="lpName">The name of the property whose value should be changed</param>
    ''' <param name="lpNewValue">The value to set</param>
    ''' <returns>True if successful, or False otherwise</returns>
    ''' <remarks>If the property scope is set to the version level, only the version specified by the lpVersionIndex will be affected.</remarks>
    Public Function ChangePropertyValue(ByVal lpName As String,
                                        ByVal lpNewValue As Object) As Boolean

      ApplicationLogging.WriteLogEntry("Enter Folder::ChangePropertyValue", TraceEventType.Verbose)

      Try


        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        If (StrComp(lpName, "FOLDER CLASS", CompareMethod.Text) = 0) OrElse (StrComp(lpName, "FOLDERCLASS", CompareMethod.Text) = 0) Then
          FolderClass = lpNewValue
        Else
          If Properties.PropertyExists(lpName) Then
            Properties(lpName).ChangePropertyValue(lpNewValue)
          Else
            Throw New PropertyDoesNotExistException(lpName)
          End If

        End If

        ApplicationLogging.WriteLogEntry("Exit Folder::ChangePropertyValue", TraceEventType.Verbose)

        Return True
        ' Change the value of the specified property

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    ''' <summary>
    ''' Gets all the properties in the collection of the specified type.
    ''' </summary>
    ''' <param name="lpPropertyType">The property type to select</param>
    ''' <returns></returns>
    ''' <remarks>If no properties exist in the collection for the specified 
    ''' type, an empty collection is returned.</remarks>
    Public Function PropertiesByType(lpPropertyType As PropertyType) As ECMProperties
      Try
        Return mobjProperties.PropertiesByType(lpPropertyType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Adds a value to the specifed property in the folder
    ''' </summary>
    ''' <param name="lpName">The name of the property whose value should be changed</param>
    ''' <param name="lpNewValue">The value to set</param>
    ''' <param name="lpAllowDuplicates">Specifies whether or not to allow duplicate values to be added.</param>
    ''' <returns>True if successful, or False otherwise</returns>
    ''' <remarks>Only valid for multi-valued properties.
    ''' 
    ''' If the property scope is set to the version level, 
    ''' only the version specified by the lpVersionIndex will be affected.</remarks>
    ''' <exception cref="Exceptions.ValueExistsException">
    ''' If lpAllowDuplicates is set to false and the value already exists, 
    ''' a ValueAlreadyExistsException will be thrown.
    ''' </exception>
    Public Function AddPropertyValue(ByVal lpName As String,
                                        ByVal lpNewValue As Object,
                                        ByVal lpAllowDuplicates As Boolean) As Boolean

      ApplicationLogging.WriteLogEntry("Enter Document::AddPropertyValue", TraceEventType.Verbose)

      Try


        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        Dim lstrErrorMessage As String = String.Empty

        If Properties(lpName).Cardinality <> Cardinality.ecmMultiValued Then
          lstrErrorMessage = String.Format("Unable to AddPropertyValue, the property '{0}' is not a multi-valued property, this action is only valid for multi-valued properties.",
                                                        lpName)
          Throw New InvalidCardinalityException(lstrErrorMessage, Cardinality.ecmMultiValued, Nothing)
          'Dim lobjDocOpEx As New InvalidOperationException(lstrErrorMessage)
          'Throw New CtsException(lstrErrorMessage, lobjDocOpEx)
        End If

        Properties(lpName).Values.Add(lpNewValue, lpAllowDuplicates)


        ApplicationLogging.WriteLogEntry("Exit Document::AddPropertyValue", TraceEventType.Verbose)

        Return True

      Catch ValueExistsEx As ValueExistsException
        ' Note it as a warning in the log and pass it on.
        ApplicationLogging.WriteLogEntry(ValueExistsEx.Message, Reflection.MethodBase.GetCurrentMethod, TraceEventType.Warning, 23976)
        Throw
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Public Function ClearPropertyValue(ByVal lpName As String) As Boolean
      Try
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        If (StrComp(lpName, "FOLDER CLASS", CompareMethod.Text) = 0) OrElse (StrComp(lpName, "FOLDERCLASS", CompareMethod.Text) = 0) Then
          FolderClass = String.Empty
        Else
          Properties(lpName).ClearPropertyValue()
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Creates a property and adds it to the folder.
    ''' </summary>
    Public Function CreateProperty(ByVal lpName As String,
                                   ByVal lpValue As Object,
                                   ByVal lpValueType As PropertyType,
                                   ByVal lpCardinality As Cardinality,
                                   ByVal lpPersistent As Boolean) As Boolean

      Try

        Dim lobjNewProperty As ECMProperty = Nothing

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        ApplicationLogging.WriteLogEntry("Enter Folder::CreateProperty", TraceEventType.Verbose)

        ' Create the new property
        lobjNewProperty = PropertyFactory.Create(lpValueType, lpName, lpName, lpCardinality, lpValue, lpPersistent)

        Properties.Add(lobjNewProperty)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ApplicationLogging.WriteLogEntry("Exit Folder::CreateProperty", TraceEventType.Verbose)
        Return False
      End Try

    End Function

    ''' <summary>
    ''' Deletes a property from the folder.
    ''' </summary>
    Public Function DeleteProperty(ByVal lpName As String) As Boolean

      Try

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        ApplicationLogging.WriteLogEntry("Enter Folder::DeleteProperty", TraceEventType.Verbose)

        If PropertyExists(lpName) Then
          Properties.Delete(lpName)
        Else
          ' The property does not exist
          ApplicationLogging.WriteLogEntry(
            String.Format("Unable to delete property '{0}': The property does not exist in the current folder.",
              lpName), TraceEventType.Warning, 2224)
          Return False
        End If

        ApplicationLogging.WriteLogEntry("Exit Folder::DeleteProperty", TraceEventType.Verbose)

        Return True

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ApplicationLogging.WriteLogEntry("Exit Folder::DeleteProperty", TraceEventType.Verbose)
        Return False
      End Try

    End Function

    ''' <summary>
    ''' Renames a property in the folder.
    ''' </summary>
    ''' <param name="lpCurrentName">The curent name of the property</param>
    ''' <param name="lpNewName">The new name for the property</param>
    ''' <param name="lpErrorMessage">A string reference to capture any error messages</param>
    ''' <returns>True if successful, otherwise false</returns>
    ''' <remarks></remarks>
    Public Function RenameProperty(ByVal lpCurrentName As String,
                                   ByVal lpNewName As String,
                                   ByRef lpErrorMessage As String) As Boolean

      Try

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        ApplicationLogging.WriteLogEntry("Enter Document::RenameProperty", TraceEventType.Verbose)

        If PropertyExists(lpCurrentName) Then
          Properties(lpCurrentName).Rename(lpNewName)
        Else
          ' The property does not exist
          lpErrorMessage = String.Format(
              "Unable to rename property '{0}' to '{1}': The property does not exist in the current folder.",
                lpCurrentName, lpNewName)
          ApplicationLogging.WriteLogEntry(lpErrorMessage, TraceEventType.Warning, 2223)
          Return False
        End If

        Return True

      Catch ex As Exception
        lpErrorMessage = ex.Message
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ApplicationLogging.WriteLogEntry("Exit Document::RenameProperty", TraceEventType.Verbose)
        Return False
      End Try

    End Function

    Public Sub Save(ByRef lpFilePath As String)
      Try

        If lpFilePath.EndsWith(DefaultFileExtension) = False Then
          lpFilePath = lpFilePath.Remove(lpFilePath.Length - 3) & DefaultFileExtension
        End If

        Serializer.Serialize.XmlFile(Me, lpFilePath)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function Transform(ByVal lpTransformation As Transformation, Optional ByRef lpErrorMessage As String = "") As Folder

      Try
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        Return lpTransformation.TransformFolder(Me, lpErrorMessage)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Public Function PropertyExists(ByVal lpName As String) As Boolean
      Try
        Return PropertyExists(lpName, True)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Determine if a property exists in the folder.
    ''' </summary>
    ''' <param name="lpName">The name of the property to check</param>
    ''' <remarks></remarks>
    Public Function PropertyExists(ByVal lpName As String, ByVal lpCaseSensitive As Boolean) As Boolean

      Try

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        Return Properties.PropertyExists(lpName, lpCaseSensitive)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return False
      End Try

    End Function


    Public Function GetSubFolderCount() As Long Implements IFolder.GetSubFolderCount
      Try
        Return 0
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetSubFolders(lpGetContents As Boolean) As Core.Folders Implements IFolder.GetSubFolders
      Try
        Return New Folders
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public ReadOnly Property HasContent(lpRecursionLevel As Core.RecursionLevel) As Boolean Implements IFolder.HasContent
      Get
        Return False
      End Get
    End Property

    Public ReadOnly Property HasSubFolders As Boolean Implements IFolder.HasSubFolders
      Get
        Return False
      End Get
    End Property

    Public Sub InitializeFolderCollection(lpFolderPath As String) Implements IFolder.InitializeFolderCollection
      Dim lstrFolderNames As String() = Nothing
      Dim lblnHasDelimiter As Boolean = True

      Try
        mstrPath = lpFolderPath
        If mstrPath.Contains("\") Then
          lstrFolderNames = mstrPath.Split("\")
        ElseIf mstrPath.Contains("/") Then
          lstrFolderNames = mstrPath.Split("/")
        Else
          lblnHasDelimiter = False
        End If

        If lblnHasDelimiter = True Then
          Name = lstrFolderNames(lstrFolderNames.GetUpperBound(0))
          If String.IsNullOrEmpty(Name) Then
            Name = mstrPath
          End If

          mstrFolders = New Collections.Specialized.StringCollection

          For lintCounter As Integer = 0 To lstrFolderNames.Length - 1
            If Not String.IsNullOrEmpty(lstrFolderNames(lintCounter)) Then
              mstrFolders.Add(lstrFolderNames(lintCounter))
            End If
          Next
        Else
          Name = mstrPath
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::CFolder::InitializeFolderCollection('{1}'", Me.GetType.Name, lpFolderPath))
      End Try
    End Sub

    Public Sub OnFolderSelected() Implements IFolder.OnFolderSelected
      ' Ignore
    End Sub

    Public Sub Refresh() Implements IFolder.Refresh
      ' Ignore
    End Sub

#End Region

#Region "Private Methods"

    Private Sub Initialize()
      Try
        AddHandler mobjProperties.CollectionChanged, AddressOf PropertiesChanged
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Function GetPathValue() As String
      Try
        Dim lobjPathProperty As ECMProperty = Nothing
        If Properties.PropertyExists("Path", False, lobjPathProperty) Then
          Return lobjPathProperty.Value
        ElseIf Properties.PropertyExists("Path Name", False, lobjPathProperty) Then
          Return lobjPathProperty.Value
        Else
          Return String.Empty
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Sub PropertiesChanged(ByVal sender As Object, e As NotifyCollectionChangedEventArgs)
      Try
        Select Case e.Action
          Case NotifyCollectionChangedAction.Add, NotifyCollectionChangedAction.Replace
            If e.NewItems.Count = 1 Then
              Dim lobjNewProperty As IProperty = e.NewItems(0)
              If lobjNewProperty IsNot Nothing Then
                Select Case lobjNewProperty.Name.ToLower()
                  Case "name"
                    Me.Name = lobjNewProperty.Value
                  Case "path"
                    mstrPath = lobjNewProperty.Value
                End Select
              End If
            End If
        End Select
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shared Function StringToASCIIKeyValuePairArray(lpString As String) As List(Of KeyValuePair)
      Try

        Dim lobjASCIIKeyList As New List(Of KeyValuePair)
        Dim lobjKeyValuePair As KeyValuePair
        For lintCounter As Integer = 0 To lpString.Length - 1
          lobjKeyValuePair = New KeyValuePair(lpString(lintCounter), Asc(lpString(lintCounter)))
          lobjASCIIKeyList.Add(lobjKeyValuePair)
        Next

        Return lobjASCIIKeyList

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
      If Not Me.disposedValue Then
        If disposing Then
          ' TODO: dispose managed state (managed objects).
        End If

        ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
        ' TODO: set large fields to null.
      End If
      Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
      ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub
#End Region

#Region "ILoggable Implementation"

    'Private mobjLogSession As Gurock.SmartInspect.Session

    'Protected Overridable Sub FinalizeLogSession() Implements ILoggable.FinalizeLogSession
    '  Try
    '    ApplicationLogging.FinalizeLogSession(mobjLogSession)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    'Protected Overridable Sub InitializeLogSession() Implements ILoggable.InitializeLogSession
    '  Try
    '    mobjLogSession = ApplicationLogging.InitializeLogSession(Me.GetType.Name, System.Drawing.Color.Goldenrod)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    'Protected Friend ReadOnly Property LogSession As Gurock.SmartInspect.Session Implements ILoggable.LogSession
    '  Get
    '    Try
    '      If mobjLogSession Is Nothing Then
    '        InitializeLogSession()
    '      End If
    '      Return mobjLogSession
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Get
    'End Property

#End Region

  End Class

End Namespace
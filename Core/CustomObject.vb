'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CustomObject.vb
'   Description :  [type_description_here]
'   Created     :  9/1/2015 11:47:40 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Collections.Specialized
Imports System.Xml.Serialization
Imports Documents.Exceptions
Imports Documents.Providers
Imports Documents.SerializationUtilities
Imports Documents.Utilities
Imports Newtonsoft.Json

#End Region

Namespace Core

  <Serializable()>
  Public Class CustomObject
    Inherits RepositoryObject
    Implements ICustomObject
    Implements IDisposable
    'Implements ILoggable

#Region "Class Constants"

    ''' <summary>
    ''' Constant value representing the file extension to save XML serialized instances
    ''' to.
    ''' </summary>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public Const CUSTOM_OBJECT_FILE_EXTENSION As String = "cof"

#End Region

#Region "Class Variables"

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private mobjAuditEvents As New AuditEvents
    <NonSerializedAttribute()>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private mobjContentSource As ContentSource
    Private mstrFolders As Collections.Specialized.StringCollection
    Private mstrLabel As String = String.Empty
    Private mstrClassName As String = String.Empty
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
        Return CUSTOM_OBJECT_FILE_EXTENSION
      End Get
    End Property

    Public Property AuditEvents As AuditEvents Implements IAuditableItem.AuditEvents
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

    ''' <summary>
    ''' Used to perform repository operations on the object.  
    ''' </summary>
    ''' <value></value>
    ''' <returns>A ContentSource object reference</returns>
    ''' <remarks></remarks>
    <XmlIgnore()>
    <JsonIgnore()>
    Public Property ContentSource As ContentSource Implements ICustomObject.ContentSource
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

    <XmlAttribute()>
    Public Property ClassName As String Implements ICustomObject.ClassName
      Get
        Try
          Return mstrClassName
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrClassName = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property Identifier As String Implements IMetaHolder.Identifier
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
    Public Property Metadata As IProperties Implements IMetaHolder.Metadata
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

    <XmlIgnore(), JsonIgnore()>
    Public Property Provider As IProvider Implements ICustomObject.Provider
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
        Dim lobjObject As CustomObject = Serializer.Deserialize.XmlFile(lpFilePath, GetType(CustomObject))
        Helper.AssignObjectProperties(lobjObject, Me)
        Initialize()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Methods"

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

        If Properties.PropertyExists(lpName) Then
          Properties(lpName).ChangePropertyValue(lpNewValue)
        Else
          Throw New PropertyDoesNotExistException(lpName)
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

        Properties(lpName).ClearPropertyValue()

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
    ''' Determine if a property exists in the object.
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
    '    mobjLogSession = ApplicationLogging.InitializeLogSession(Me.GetType.Name, System.Drawing.Color.Coral)
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
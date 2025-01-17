'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Collections.ObjectModel
Imports System.Xml
Imports System.Xml.Serialization
Imports Documents.Annotations.Auditing
Imports Documents.Annotations.Common
Imports Documents.Annotations.Highlight
Imports Documents.Annotations.Redaction
Imports Documents.Annotations.Reference
Imports Documents.Annotations.Security
Imports Documents.Annotations.Shape
Imports Documents.Annotations.Special
Imports Documents.Annotations.Text
Imports Documents.Core
Imports Documents.SerializationUtilities
Imports Documents.Utilities

#End Region

'todo: add constructors for filepath and xmlstring
Namespace Annotations
  ''' <summary>
  ''' Base class for describing an annotation's security, display, layout and composition.
  ''' </summary>
  <Serializable()>
  <XmlInclude(GetType(TextAnnotation))>
  <XmlInclude(GetType(LineBase))>
  <XmlInclude(GetType(SpecialBase))>
  <XmlInclude(GetType(ShapeBase))>
  <XmlInclude(GetType(RedactionBase))>
  <XmlInclude(GetType(ReferenceBase))>
  <XmlInclude(GetType(HighlightBase))>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class Annotation
    Implements IDisposable
    Implements ISerialize

#Region "Class Constants"

    ''' <summary>
    ''' Constant value representing the file extension to save XML serialized instances
    ''' to.
    ''' </summary>
    Public Const ANNOTATION_FILE_EXTENSION As String = "caf"

#End Region

#Region "Class Variables"

    Private Const m_defaultFileExtension As String = "caf"
    'Private m_Composite As Annotations
    Private m_objProperties As New ECMProperties
    'Protected m_Parent As Annotation
    'Protected m_AnnotatedContent As Content

    Public Property Annotations As Annotations

#End Region

#Region "Public Properties"


    <XmlAttribute("Dpi")>
    Public Property Dpi As Single

    ''' <summary>
    ''' Gets or sets the ID.
    ''' </summary>
    ''' <value>The ID.</value>
    <XmlIgnore()>
    Public Property ID() As String
      Get
        If Me.Properties.Contains("ID") Then
          Return Me.Properties("ID").Value
        End If

        Return String.Empty
      End Get
      Set(ByVal value As String)
        If Me.Properties.Contains("ID") Then
          Me.Properties("ID").Value = value
        Else
          'Dim idProperty As New ECMProperty(PropertyType.ecmString, "ID", Cardinality.ecmSingleValued, String.Empty, True)
          'idProperty.Value = value
          Dim idProperty As SingletonStringProperty = PropertyFactory.Create(PropertyType.ecmString, "ID", value)
          'idProperty.Value = value
          Me.Properties.Add(idProperty)
        End If
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the annotated content id.
    ''' This is useful for communicating the document and content element to the rest of the provider.
    ''' </summary>
    ''' <value>The annotated content id.</value>
    <XmlIgnore()>
    Public Property AnnotatedContentId() As String

    ''' <summary>
    ''' Gets or sets the name of the class.
    ''' Used to specify a specific annotation subclass of the platform, similar to document class.
    ''' This property may remove the need to persist security.
    ''' </summary>
    ''' <value>The name of the class.</value>
    <System.Xml.Serialization.XmlAttribute("Class")>
    Public Property ClassName() As String

    ''' <summary>
    ''' Gets or sets the access controls for this annotation.
    ''' </summary>
    ''' <value>The access controls for this annotation.</value>
    Public Property AccessControls() As Collection(Of AnnotationAccessControl)

    ''' <summary>
    ''' Gets or sets the audit events.
    ''' </summary>
    ''' <value>The audit events.</value>
    Public Property AuditEvents() As AuditRecord

    ''' <summary>
    ''' Gets or sets the display settings.
    ''' </summary>
    ''' <value>The display settings.</value>
    Public Property Display() As DisplaySettings

    ''' <summary>
    ''' Gets or sets the layout settings.
    ''' </summary>
    ''' <value>The layout settings.</value>
    Public Property Layout() As LayoutSettings

    ''' <summary>
    ''' Gets or sets the collection of properties for the document.
    ''' </summary>
    Public Property Properties() As ECMProperties
      Get

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        Return m_objProperties

      End Get
      Set(ByVal value As ECMProperties)

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        m_objProperties = value

      End Set
    End Property

    ''' <summary>
    ''' Gets or sets a value indicating whether this instance is redaction.
    ''' </summary>
    ''' <value>
    ''' <c>true</c> if this instance is redaction; otherwise, <c>false</c>.
    ''' </value>
    <System.Xml.Serialization.XmlAttribute()>
    Property IsRedaction() As Boolean

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
        Return ANNOTATION_FILE_EXTENSION
      End Get
    End Property

    'Public ReadOnly Property Parent As Annotation
    '  Get
    '    Return m_Parent
    '  End Get
    'End Property

    <XmlIgnore()>
    Public Property AnnotatedContent As Content
#End Region

#Region "Private Properties"

    Private ReadOnly Property IsDisposed() As Boolean
      Get
        Return disposedValue
      End Get
    End Property

#End Region

#Region "Private Methods"

    Protected Overridable Function DebuggerIdentifier() As String
      Try
        If ID Is Nothing OrElse ID.Length = 0 Then
          Return String.Format("{0}  ({1})", ClassName, AuditEvents.DebuggerIdentifier)
        Else
          Return String.Format("{0}: {1} ({2})", ID, ClassName, AuditEvents.DebuggerIdentifier)
        End If
      Catch ex As System.Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Initializes a new instance of the <see cref="Annotation"/> class.
    ''' </summary>
    Public Sub New()
      Try
        Me.Properties = New ECMProperties()
        Me.ID = String.Empty
        Me.ClassName = "Annotation"
        Me.IsRedaction = False
        Me.AccessControls = New Collection(Of AnnotationAccessControl)()
        Me.AuditEvents = New AuditRecord()
        'Me.Composite = New Annotations()
        Me.Display = New DisplaySettings()
        Me.Layout = New LayoutSettings()
      Catch ex As System.Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the <see cref="Annotation" /> class from a content element.
    ''' Preserves the overall content element associated with the annotations.
    ''' </summary>
    ''' <param name="contentElement">The content element.</param>
    Public Sub New(ByVal contentElement As Content)
      Try
        Me.Properties = New ECMProperties()
        Me.ID = String.Empty
        Me.ClassName = "Annotation"
        Me.IsRedaction = False
        Me.AccessControls = New Collection(Of AnnotationAccessControl)()
        Me.AuditEvents = New AuditRecord()
        'Me.Composite = New Annotations()
        Me.Display = New DisplaySettings()
        Me.Layout = New LayoutSettings()

        Me.AnnotatedContent = contentElement

      Catch ex As System.Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "IDisposable Support"

    Private disposedValue As Boolean     ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
      Try
        If Not Me.disposedValue Then
          If disposing Then
            ' TODO: free other state (managed objects).
            'If m_Composite IsNot Nothing Then
            '  For Each item As Annotation In Composite ' Composite.MemberSet
            '    item.Dispose()
            '  Next
            'End If
          End If

          Me.AccessControls = Nothing
          Me.AuditEvents = Nothing
          Me.ClassName = Nothing
          Me.Display = Nothing
          Me.Layout = Nothing
          Me.m_objProperties = Nothing

        End If
        Me.disposedValue = True
      Catch ex As System.Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
      ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub

#End Region

#Region "ISerialize Support"

    ''' <summary>
    ''' Serializes the specified lp file path.
    ''' </summary>
    ''' <param name="lpFilePath">The lp file path.</param>
    Public Sub Serialize(ByVal lpFilePath As String) Implements ISerialize.Serialize
      Try
        Serialize(lpFilePath, "")
      Catch ex As System.Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Serializes the specified lp file path.
    ''' </summary>
    ''' <param name="lpFilePath">The lp file path.</param>
    ''' <param name="lpFileExtension">The lp file extension.</param>
    Public Sub Serialize(ByRef lpFilePath As String, ByVal lpFileExtension As String) Implements ISerialize.Serialize
      Serializer.Serialize.XmlFile(Me, lpFilePath)
    End Sub

    ''' <summary>
    ''' Serializes the specified lp file path.
    ''' </summary>
    ''' <param name="lpFilePath">The lp file path.</param>
    ''' <param name="lpWriteProcessingInstruction">if set to <c>true</c> [lp write processing instruction].</param>
    ''' <param name="lpStyleSheetPath">The lp style sheet path.</param>
    Public Sub Serialize(ByVal lpFilePath As String, ByVal lpWriteProcessingInstruction As Boolean, ByVal lpStyleSheetPath As String) Implements ISerialize.Serialize

      If lpWriteProcessingInstruction = True Then
        Serializer.Serialize.XmlFile(Me, lpFilePath, , , True, lpStyleSheetPath)
      Else
        Serializer.Serialize.XmlFile(Me, lpFilePath)
      End If

    End Sub

    ''' <summary>
    ''' Serializes this instance.
    ''' </summary>
    ''' <returns></returns>
    Public Function Serialize() As System.Xml.XmlDocument Implements ISerialize.Serialize
      Return Serializer.Serialize.Xml(Me)
    End Function

    Public Function ToXmlString() As String Implements ISerialize.ToXmlString

      Try
        Return Serializer.Serialize.XmlString(Me)
      Catch ex As System.Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    ''' <summary>
    ''' Deserializes the specified lp file path.
    ''' </summary>
    ''' <param name="lpFilePath">The lp file path.</param>
    ''' <param name="lpErrorMessage">The lp error message.</param>
    ''' <returns></returns>
    Public Function Deserialize(ByVal lpFilePath As String, Optional ByRef lpErrorMessage As String = "") As Object Implements ISerialize.Deserialize
      Try
        Return Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType)
      Catch ex As System.Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        lpErrorMessage = Helper.FormatCallStack(ex)
        Return Nothing
      End Try
    End Function

    ''' <summary>
    ''' Deserializes the specified lp XML.
    ''' </summary>
    ''' <param name="lpXML">The lp XML.</param>
    ''' <returns></returns>
    Public Function Deserialize(ByVal lpXML As XmlDocument) As Object Implements ISerialize.Deserialize
      Try

        Dim lobjAnnotation As Annotation = Serializer.Deserialize.XmlString(lpXML.OuterXml, Me.GetType)

        If lobjAnnotation IsNot Nothing Then
          Return lobjAnnotation
        Else
          Return Nothing
        End If

      Catch ex As System.Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Helper.DumpException(ex)
        ' Rethrow the exception
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace

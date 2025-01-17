'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"


Imports System.Security
Imports System.Xml.Serialization
Imports Documents.Security
Imports Documents.Utilities
Imports Newtonsoft.Json

#End Region

Namespace Core

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class Version
    Implements IAuditableItem
    Implements IMetaHolder
    Implements IDisposable

#Region "Class Variables"

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private mintID As Integer

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private mobjAuditEvents As New AuditEvents

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private mobjProperties As New ECMProperties
    'Private mstrContentPath As String

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private mobjPermissions As IPermissions = New Permissions

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private WithEvents MobjContents As New Contents(Me)

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private mobjDocument As Document

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private mobjRelationships As New Relationships

#End Region

#Region "Public Properties"

#Region "IMetaHolder Implementation"

    <JsonIgnore()>
    Public Property Identifier() As String Implements IMetaHolder.Identifier
      Get
        Return ID
      End Get
      Set(ByVal value As String)
        If Helper.IsNumeric(value) = False Then
          Throw New ArgumentException(String.Format("The value of '{0}' is not a valid identifier for a Version object.  The value must be numeric.", value), NameOf(value))
        Else
          ID = value
        End If
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the collection of metadata properties for the object.
    ''' </summary>
    ''' <returns>ECMProperties collection</returns>
    ''' <remarks>Interface passthrough to the Properties collection.</remarks>
    <XmlIgnore()>
    <JsonIgnore()>
    Public Property Metadata() As IProperties Implements IMetaHolder.Metadata
      Get
        Return Properties
      End Get
      Set(ByVal value As IProperties)
        Properties = value
      End Set
    End Property

#End Region

#Region "Read Only Properties"

    ''' <summary>
    ''' Gets the total number of content elements for the version.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonIgnore()>
    Public ReadOnly Property ContentCount As Integer
      Get
        Try
          Return Contents.Count
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    <JsonIgnore()>
    Public ReadOnly Property Document() As Document
      Get
        Return mobjDocument
      End Get
    End Property


#If CTSCORE = 1 Then
    ''' <summary>
    ''' Gets the first content element if available.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonIgnore()> _
    Public ReadOnly Property PrimaryContent As Content
      Get
        Return GetPrimaryContent()
      End Get
    End Property
#Else
    ''' <summary>
    ''' Gets the first content element if available.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property PrimaryContent As Content
      Get
        Return GetPrimaryContent()
      End Get
    End Property

#End If
    ''' <summary>
    ''' Gets the total content size of all content elements combined.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property TotalContentSize As FileSize
      Get
        Try
          Return GetTotalContentSize()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

    <XmlAttribute()>
    Public Property ID() As Integer
      Get
        Return mintID
      End Get
      Set(ByVal value As Integer)
        mintID = value
      End Set
    End Property

    Public Property Contents() As Contents
      Get
        Return MobjContents
      End Get
      Set(ByVal Value As Contents)
        MobjContents = Value
      End Set
    End Property

#Region "IAuditableItem Implementation"

    Public Property AuditEvents As AuditEvents Implements IAuditableItem.AuditEvents
      Get
        Return mobjAuditEvents
      End Get
      Set(value As AuditEvents)
        mobjAuditEvents = value
      End Set
    End Property

#End Region

    ''' <summary>
    ''' Gets or sets the parent and/or child relationships for the version.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Relationships() As Relationships
      Get

        Return mobjRelationships

      End Get
      Set(ByVal value As Relationships)

        mobjRelationships = value

      End Set
    End Property

    Public Property Properties() As ECMProperties
      Get
        Return mobjProperties
      End Get
      Set(ByVal value As ECMProperties)
        mobjProperties = value
      End Set
    End Property

    ''' <summary>
    ''' Gets the permissions for the version
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Permissions As IPermissions
      Get
        Return mobjPermissions
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal lpDocument As Document)
      MyBase.New()
      Try
        mobjDocument = lpDocument
        ' Default the ID using the version count for the parent document
        ' The Version ID should correspond to its index number in the 
        ' Versions collection (zero based)
        ID = lpDocument.Versions.Count
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    'Public Sub New(ByVal lpDocument As Document)
    '  MyBase.New()
    '  mobjDocument = lpDocument
    'End Sub

    Public Sub New(ByVal lpVersionID As String)
      MyBase.New()
      Me.ID = lpVersionID
    End Sub

    'Public Sub New(ByVal lpDocument As Document, ByVal lpVersionID As String)
    '  MyBase.New()
    '  Me.ID = lpVersionID
    '  mobjDocument = lpDocument
    'End Sub

    Public Sub New(ByVal lpVersionID As String, ByVal lpProperties As ECMProperties)
      MyBase.New()
      Me.ID = lpVersionID
      Me.Properties = lpProperties
    End Sub

    'Public Sub New(ByVal lpDocument As Document, ByVal lpVersionID As String, ByVal lpProperties As ECMProperties)
    '  MyBase.New()
    '  Me.ID = lpVersionID
    '  Me.Properties = lpProperties
    '  mobjDocument = lpDocument
    'End Sub

    Public Sub New(ByVal lpVersionID As String, ByVal lpProperties As ECMProperties, ByVal lpContentPath As String)
      MyBase.New()
      Me.ID = lpVersionID
      Me.Properties = lpProperties
      Me.Contents.Add(lpContentPath)
    End Sub

    'Public Sub New(ByVal lpDocument As Document, ByVal lpVersionID As String, ByVal lpProperties As ECMProperties, ByVal lpContentPath As String)
    '  MyBase.New()
    '  Me.ID = lpVersionID
    '  Me.Properties = lpProperties
    '  Me.Contents.Add(lpContentPath)
    '  mobjDocument = lpDocument
    'End Sub

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Gets all the properties in the collection of the specified type.
    ''' </summary>
    ''' <param name="lpPropertyType">The property type to select</param>
    ''' <returns></returns>
    ''' <remarks>If no properties exist in the collection for the specified 
    ''' type, an empty collection is returned.</remarks>
    Public Function PropertiesByType(ByVal lpPropertyType As PropertyType) As ECMProperties
      Try
        Return mobjProperties.PropertiesByType(lpPropertyType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Returns the first content element of the version if available
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>If there is no content in this version a DocumentException will be thrown.</remarks>
    Public Function GetPrimaryContent() As Content
      Try
        If HasContent() = False Then
          Throw New Exceptions.DocumentException(Document, ID, "This version has no content")
        End If

        ' Use the enumerator to return the first content element added
        For Each lobjContent As Content In Contents
          Return lobjContent
        Next

        ' This should never get called but it satisfies the 
        ' compiler warning for potentially never returning a value
        Return Nothing

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      Finally

      End Try
    End Function

    ''' <summary>
    ''' Determines whether or not the version contains any content elements
    ''' </summary>
    ''' <returns>True if yes or False if no</returns>
    ''' <remarks></remarks>
    Public Function HasContent() As Boolean
      Try
        If Contents IsNot Nothing AndAlso Contents.Count > 0 Then
          Return True
        Else
          Return False
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#Region "ShouldSerialize Methods"

    Public Function ShouldSerializeAuditEvents() As Boolean
      Try
        Return AuditEvents.Count > 0
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ShouldSerializeRelationships() As Boolean
      Try
        Return Relationships.Count > 0
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ShouldSerializePermissions() As Boolean
      Try
        Return Permissions.Count > 0
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#End Region

#Region "Friend Methods"

    Friend Sub SetDocument(ByVal lpDocument As Document)
      Try
        mobjDocument = lpDocument
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

    Protected Friend Function DebuggerIdentifier() As String
      Try

        If disposedValue = True Then
          Return "Version Disposed"
        End If

        If Contents Is Nothing OrElse Contents.Count = 0 Then
          Return String.Format("{0}: ({1} Properties) {2}", ID, Properties.Count, "No Content")
        ElseIf Contents.Count > 1 Then
          Return String.Format("{0}: ({1} Properties) {2} + {3} other content elements.",
                               ID, Properties.Count, TotalContentSize.ToString, Contents.Count - 1)
        Else
          Return String.Format("{0}: ({1} Properties) {2}", ID, Properties.Count, GetPrimaryContent.DebuggerIdentifier)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Deletes all the files referenced by all the contents via ContentPath (if they exist).
    ''' </summary>
    Friend Sub DeleteContentFiles()
      Try
        If Me.Contents IsNot Nothing Then
          For Each lobjContent As Content In Me.Contents
            lobjContent.DeleteContentFiles()
          Next
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Function GetTotalContentSize() As FileSize
      Try
        Dim llngTotalByteCount As Long = 0
        If HasContent() Then
          For Each lobjContent As Content In Me.Contents
            llngTotalByteCount += lobjContent.FileSize.Bytes
          Next
        End If
        Return New FileSize(llngTotalByteCount)
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
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
      If Not Me.disposedValue Then
        If disposing Then
          ' DISPOSETODO: dispose managed state (managed objects).

          ' Dispose of all properties
          Me.Properties.Dispose()

          ' Dispose of all contents
          Me.Contents.Dispose()

          Me.Properties = Nothing
          Me.Contents = Nothing

        End If

        ' DISPOSETODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
        ' DISPOSETODO: set large fields to null.
      End If
      Me.disposedValue = True
    End Sub

    ' DISPOSETODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
      ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub
#End Region

  End Class

End Namespace
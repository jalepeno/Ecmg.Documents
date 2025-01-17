'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Providers
Imports Documents.SerializationUtilities
Imports Documents.Utilities

Namespace Core
  ''' <summary>Represents a single object contained in a folder.</summary>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class FolderContent
    Implements IComparable

#Region "Class Variables"

    Private mobjProperties As New FolderContentProperties
    Private mobjParentFolder As IFolder
    Private mstrParentFolderId As String = String.Empty
    'Private mlngContentSize As Long
    'Private mstrFileName As String
    'Private mobjDateLastModified As DateTime

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpContentName As Object)
      Me.New(lpContentName.ToString)
    End Sub

    Public Sub New(ByVal lpContentName As String)
      Properties.Add("Name", lpContentName)
    End Sub

    Public Sub New(ByVal lpContentName As String, ByVal lpId As String)
      Properties.Add("Name", lpContentName)
      Properties.Add("ID", lpId)
    End Sub

    Public Sub New(ByVal lpParentFolder As IFolder)
      mobjParentFolder = lpParentFolder
    End Sub

    Public Sub New(ByVal lpContentName As String, ByVal lpParentFolder As IFolder)
      Me.New(lpParentFolder)
      Properties.Add("Name", lpContentName)
    End Sub

    Public Sub New(ByVal lpContentName As String, ByVal lpId As String, ByVal lpParentFolder As IFolder)
      Me.New(lpParentFolder)
      Properties.Add("Name", lpContentName)
      Properties.Add("ID", lpId)
    End Sub

    Public Sub New(ByVal lpContentName As String,
                   ByVal lpId As String,
                   ByVal lpSize As Long,
                   ByVal lpFileName As String,
                   ByVal lpLastModifiedDate As DateTime)
      Properties.Add("Name", lpContentName)
      Properties.Add("ID", lpId)
      Properties.Add("ContentSize", lpSize)
      Properties.Add("FileName", lpFileName)
      Properties.Add("DateLastModified", lpLastModifiedDate)
    End Sub

#End Region

#Region "Public Properties"

    Public ReadOnly Property Size() As Long
      Get
        Try
          If Not Properties.Item("ContentSize") Is Nothing Then
            Return Properties.Item("ContentSize")
          Else
            Return 0
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          Return 0
        End Try
      End Get
    End Property

    Public ReadOnly Property FileName() As String
      Get
        Try
          If Not Properties.Item("FileName") Is Nothing Then
            Return Properties.Item("FileName")
          Else
            Return ""
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          Return ""
        End Try
      End Get
    End Property

    Public ReadOnly Property LastModifiedDate() As DateTime
      Get
        Try
          If Not Properties.Item("DateLastModified") Is Nothing Then
            Return Properties.Item("DateLastModified")
          ElseIf Not Properties.Item("Date Modified") Is Nothing Then
            Return Properties.Item("Date Modified")
          Else
            Return Nothing
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          Return Nothing
        End Try
      End Get
    End Property

    ''' <summary>
    ''' Gets a value indicating whether or not the item 
    ''' is a product of a previous export operation.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property IsExportBased As Boolean
      Get
        Try
          Return DetermineIfIsExportBased()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property Properties() As FolderContentProperties
      Get
        Return mobjProperties
      End Get
    End Property

    Public ReadOnly Property Name() As String
      Get
        Try
          If Not Properties.Item("Name") Is Nothing Then
            Return Properties.Item("Name")
          Else
            Return ""
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          Return ""
        End Try
      End Get
    End Property

    Public ReadOnly Property ID() As String
      Get
        Try
          If Not Properties.Item("ID") Is Nothing Then
            Return Properties.Item("ID")
          ElseIf Not Properties.Item("Id") Is Nothing Then
            Return Properties.Item("Id")
          ElseIf Not Properties.Item("id") Is Nothing Then
            Return Properties.Item("id")
          ElseIf Not Properties.Item("Name") Is Nothing Then
            Return Properties.Item("Name")
          Else
            Return ""
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          Return ""
        End Try
      End Get
    End Property

    Public ReadOnly Property ParentFolder() As IFolder
      Get
        Return mobjParentFolder
      End Get
    End Property

    ''' <summary>
    ''' This gets/sets the ID of the folder in which this content is living
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ParentFolderId() As String
      Get
        Return mstrParentFolderId
      End Get
      Set(ByVal value As String)
        mstrParentFolderId = value
      End Set
    End Property

#End Region

#Region "Public Methods"

#End Region

#Region "Private Methods"

    Private Function DebuggerIdentifier() As String
      Try
        Return Name
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Determines if the item is the product of a previous export operation.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function DetermineIfIsExportBased() As Boolean
      Try

        If String.IsNullOrEmpty(FileName) Then
          Return False
        End If

        If FileName.EndsWith(".cpf") OrElse FileName.EndsWith(".cdf") Then
          Return True
        End If

        Return False

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Public Overridable Methods"

    Public Function ToXmlString() As String
      Try
        Return Serializer.Serialize.XmlString(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return Me.GetType.FullName
      End Try
    End Function

#End Region

#Region "IComparable Implementation"

    Public Function CompareTo(ByVal obj As Object) As Integer Implements IComparable.CompareTo

      Try
        If TypeOf obj Is FolderContent Then
          Return Name.CompareTo(obj.Name)
        Else
          Throw New ArgumentException("Object is not a FolderContent Object")
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("FolderContent::CompareTo('{0}')", obj.ToString))
      End Try

    End Function

#End Region

#Region "Child Classes"

    Public Class FolderContentProperties
      Inherits Collections.Specialized.NameValueCollection
      Implements ICloneable

      Public Shadows Function Item(ByVal lpName As String) As Object
        Try
          'Dim lobjArrayList As ArrayList = Nothing
          Dim lobjReturnValue As Object = MyBase.BaseGet(lpName)
          If lobjReturnValue.GetType.Name = "ArrayList" Then
            If lobjReturnValue.Count > 0 Then
              Return lobjReturnValue(0)
            Else
              Return ""
            End If
          Else
            Return lobjReturnValue
          End If
          'lobjArrayList = MyBase.BaseGet(lpName)
          'If lobjArrayList.Count > 0 Then
          '  Return lobjArrayList(0)
          'Else
          '  Return ""
          'End If
        Catch ex As Exception
          ' Just return an empty string
          Return ""
        End Try
      End Function

      Public Overloads Sub Add(ByVal name As String, ByVal value As Object)
        MyBase.BaseAdd(name, value)
      End Sub

      Public Function Clone() As Object Implements System.ICloneable.Clone

        Try

          Dim lobjFolderContentProperties As New FolderContentProperties

          For Each lstrKey As String In Me.Keys
            lobjFolderContentProperties.Add(lstrKey, Item(lstrKey))
          Next

          Return lobjFolderContentProperties

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Function

    End Class

#End Region

  End Class

End Namespace
'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------


Imports Documents.Arguments
Imports Documents.Core
Imports Documents.Utilities

Namespace Providers
  ''' <summary>Base folder class from which all provider specific folder objects inherit.</summary>
  <Serializable()>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public MustInherit Class CFolder
    Inherits RepositoryObject
    Implements IFolder
    Implements IComparable

#Region "Public Events"

    Public Event FolderSelected(ByVal sender As Object, ByVal e As Arguments.FolderEventArgs) _
    Implements IFolder.FolderSelected

#End Region

#Region "Class Constants"

    Public Const ALL_CONTENT As Long = -1
    Public Const NOT_YET_INITIALIZED As Long = -100

#End Region

#Region "Class Variables"

    Protected mstrPath As String = String.Empty
    Private mstrLabel As String = String.Empty
    Protected mobjSubFolders As New Folders
    Protected mstrFolders As Collections.Specialized.StringCollection
    Private mobjProvider As IProvider
    Private mblnInvisiblePassThrough As Boolean
    Private mlngContentCount As Long = NOT_YET_INITIALIZED
    Private mobjContentSource As ContentSource
    Private mlngMaxContentCount As Long

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private mobjAuditEvents As New AuditEvents

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpFolderPath As String, ByVal lpMaxContentCount As Long)
      InitializeFolderCollection(lpFolderPath)
      MaxContentCount = lpMaxContentCount
    End Sub

    Public Sub New(ByRef lpProvider As IProvider,
                   Optional ByVal lpContentSource As ContentSource = Nothing)

      Provider = lpProvider
      ContentSource = lpContentSource
      mobjSubFolders = New Folders(lpProvider)
      InitializeFolder()

    End Sub

    Public Sub New(ByVal lpFolderPath As String,
                   ByVal lpProvider As IProvider,
                   Optional ByVal lpContentSource As ContentSource = Nothing)

      InitializeFolderCollection(lpFolderPath)
      Provider = lpProvider
      ContentSource = lpContentSource
      mobjSubFolders = New Folders(lpProvider)
      InitializeFolder()

    End Sub

#End Region

    Public Sub InitializeFolderCollection(ByVal lpFolderPath As String) Implements IFolder.InitializeFolderCollection

      Dim lstrFolderNames As String() = Nothing
      Dim lblnHasDelimiter As Boolean = True

      Try
        mstrPath = lpFolderPath
        If mstrPath.Contains("\"c) Then
          lstrFolderNames = mstrPath.Split("\")
        ElseIf mstrPath.Contains("/"c) Then
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

#Region "Public Properties"

    Public Overridable Property Label() As String Implements IFolder.Label
      Get
        Return mstrLabel
      End Get
      Set(ByVal value As String)
        mstrLabel = value
        AddProperty(Reflection.MethodBase.GetCurrentMethod, value)
      End Set
    End Property

    Public Property Identifier() As String Implements IMetaHolder.Identifier
      Get
        Return Me.Id
      End Get
      Set(value As String)
        Me.Id = value
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

    Public Overloads Property Metadata As IProperties Implements IFolder.Metadata
      Get
        Return Me.Properties
      End Get
      Set(value As IProperties)
        Me.Properties = value
      End Set
    End Property

    Public Overridable ReadOnly Property TreeLabel() As String Implements IFolder.TreeLabel
      Get
        If mstrLabel.Length = 0 Then
          Return Name
        Else
          Return Label & " (" & Name & ")"
        End If
      End Get
    End Property

    ''' <summary>
    ''' Gets or Sets the maximum number of contents to be retrieved for the folder
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property MaxContentCount() As Long Implements IFolder.MaxContentCount
      Get
        Return mlngMaxContentCount
      End Get
      Set(ByVal value As Long)
        mlngMaxContentCount = value
      End Set
    End Property

    Public ReadOnly Property HasContent(ByVal lpRecursionLevel As Core.RecursionLevel) As Boolean Implements IFolder.HasContent
      Get
        Try
          Select Case lpRecursionLevel
            Case RecursionLevel.ecmThisLevelOnly
              If Me.Contents.Count > 0 Then
                Return True
              Else
                Return False
              End If
            Case RecursionLevel.ecmThisLevelAndImmediateChildrenOnly
              If HasContent(RecursionLevel.ecmThisLevelOnly) = True Then
                Return True
              Else
                If Me.SubFolderCount > 0 Then
                  For Each lobjfolder As CFolder In Me.SubFolders
                    If lobjfolder.HasContent(RecursionLevel.ecmThisLevelOnly) = True Then
                      Return True
                    End If
                  Next
                End If
              End If
              Return False
            Case RecursionLevel.ecmAllChildren
              If HasContent(RecursionLevel.ecmThisLevelAndImmediateChildrenOnly) = True Then
                Return True
              Else
                For Each lobjFolder As CFolder In Me.SubFolders
                  If lobjFolder.HasContent(RecursionLevel.ecmThisLevelAndImmediateChildrenOnly) = True Then
                    Return True
                  Else
                    Return lobjFolder.HasContent(RecursionLevel.ecmAllChildren)
                  End If
                Next
              End If
          End Select
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Overridable ReadOnly Property Path() As String Implements IFolder.Path
      Get
        If mstrPath <> String.Empty Then
          Return mstrPath
        Else
          mstrPath = GetPath()
          AddProperty(Reflection.MethodBase.GetCurrentMethod, mstrPath)
          Return mstrPath
        End If
      End Get
    End Property

    Public Property ContentSource() As ContentSource Implements IFolder.ContentSource
      Get
        Return mobjContentSource
      End Get
      Set(ByVal value As ContentSource)
        mobjContentSource = value
        AddProperty(Reflection.MethodBase.GetCurrentMethod, value, PropertyType.ecmObject)
      End Set
    End Property

    Public Property Provider() As IProvider Implements IFolder.Provider
      Get
        Return mobjProvider
      End Get
      Set(ByVal value As IProvider)
        mobjProvider = value
        AddProperty(Reflection.MethodBase.GetCurrentMethod, value, PropertyType.ecmObject)
      End Set
    End Property

    Public Property InvisiblePassThrough() As Boolean Implements IFolder.InvisiblePassThrough
      Get
        Return mblnInvisiblePassThrough
      End Get
      Set(ByVal value As Boolean)
        mblnInvisiblePassThrough = value
        AddProperty(Reflection.MethodBase.GetCurrentMethod, value, PropertyType.ecmBoolean)
      End Set
    End Property

    Public ReadOnly Property HasSubFolders() As Boolean Implements IFolder.HasSubFolders
      Get
        If SubFolderCount > 0 Then
          AddProperty(Reflection.MethodBase.GetCurrentMethod, True, PropertyType.ecmBoolean)
          Return True
        Else
          AddProperty(Reflection.MethodBase.GetCurrentMethod, False, PropertyType.ecmBoolean)
          Return False
        End If
      End Get
    End Property

    Public ReadOnly Property SubFolderCount() As Long Implements IFolder.SubFolderCount
      Get
        Try
          Dim llngSubFolderCount As Long = GetSubFolderCount()
          AddProperty(Reflection.MethodBase.GetCurrentMethod, llngSubFolderCount, PropertyType.ecmLong)
          Return llngSubFolderCount
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("{0}::CFolder::SubFolderCount", Me.GetType.Name))
        End Try
      End Get
    End Property

    Public Property FolderName As String Implements IFolder.Name
      Get
        Return Name
      End Get
      Set(ByVal value As String)
        Name = value
      End Set
    End Property

    Public ReadOnly Property FolderNames() As Collections.Specialized.StringCollection Implements IFolder.FolderNames
      Get
        Return mstrFolders
      End Get
    End Property

    Public ReadOnly Property ContentCount() As Long Implements IFolder.ContentCount
      Get
        Try
          Dim llngContentCount As Long = ContentCount(RecursionLevel.ecmAllChildren)
          Return llngContentCount
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("{0}::CFolder::Get_ContentCount", Me.GetType.Name))
        End Try
      End Get
    End Property

    Public ReadOnly Property ContentCount(ByVal FolderRecursionLevel As Core.RecursionLevel) As Long Implements IFolder.ContentCount
      Get
        Try
          If mlngContentCount = NOT_YET_INITIALIZED Then
            mlngContentCount = GetContentCount(FolderRecursionLevel)
          End If
          Return mlngContentCount
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("{0}::CFolder::Get_ContentCount('{1}')", Me.GetType.Name, FolderRecursionLevel.ToString))
        End Try
      End Get
    End Property

    ''' <summary>
    ''' SubFolders
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>This method should retrieve subfolders and it's contents</remarks>
    Public Overridable ReadOnly Property SubFolders() As Folders Implements IFolder.SubFolders
      Get
        Return SubFolders(True)
      End Get
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="lpGetContents"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>In some cases you want to retrieve the contents of subfolders (i.e. an export) and in other cases you do not (i.e. folder browsing)</remarks>
    Public Overridable ReadOnly Property SubFolders(ByVal lpGetContents As Boolean) As Folders
      Get
        Try
          Return GetSubFolders(lpGetContents)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ''' <summary>
    ''' SubFolders
    ''' </summary>
    ''' <param name="lpGetContents"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>In some cases you want to retrieve the contents of subfolders (i.e. an export) and in other cases you do not (i.e. folder browsing)</remarks>
    Public MustOverride Function GetSubFolders(ByVal lpGetContents As Boolean) As Folders Implements IFolder.GetSubFolders

    Public MustOverride ReadOnly Property Contents() As FolderContents Implements IFolder.Contents

#End Region

#Region "Protected MustOverride Methods"

    Public MustOverride Function GetID() As String Implements IFolder.GetID

    Protected MustOverride Function GetPath() As String Implements IFolder.GetPath

    Protected MustOverride Function GetSubFolderCount() As Long Implements IFolder.GetSubFolderCount

    Protected MustOverride Function GetFolderByPath(ByVal lpFolderPath As String, ByVal lpMaxContentCount As Long) As IFolder Implements IFolder.GetFolderByPath

    Protected MustOverride Sub InitializeFolder()

    Public MustOverride Sub Refresh() Implements IFolder.Refresh

#End Region

#Region "Public Methods"

    Public Function GetContentCount() As Long Implements IFolder.GetContentCount
      Return GetContentCount(RecursionLevel.ecmThisLevelOnly)
    End Function

    Public Function GetContentCount(ByVal lpFolderRecursionLevel As _
                                     Core.RecursionLevel) As Long Implements IFolder.GetContentCount

      Try
        Dim llngContentCount As Long = 0

        ' Get the Content count at this level
        If Contents IsNot Nothing Then
          llngContentCount += Me.Contents.Count
        Else
          'llngContentCount = 0
        End If

        If lpFolderRecursionLevel <> RecursionLevel.ecmThisLevelOnly AndAlso Me.HasSubFolders = True Then
          ' We have subfolders and we want to include them in the count
          Select Case lpFolderRecursionLevel
            Case RecursionLevel.ecmThisLevelAndImmediateChildrenOnly
              For Each lobjSubFolder As IFolder In Me.SubFolders
                llngContentCount += GetContentCount(RecursionLevel.ecmThisLevelOnly)
              Next
            Case RecursionLevel.ecmAllChildren
              For Each lobjSubFolder As IFolder In Me.SubFolders
                llngContentCount += lobjSubFolder.GetContentCount(RecursionLevel.ecmAllChildren)
              Next
          End Select
        End If

        Return llngContentCount
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::CFolder({1})::GetContentCount('{2}')", Me.GetType.Name, Me.Path, lpFolderRecursionLevel.ToString))
      End Try

    End Function

    Public Sub OnFolderSelected() Implements IFolder.OnFolderSelected
      Try
        RaiseEvent FolderSelected(Me, New FolderEventArgs(Me))
        Me.ContentSource.SelectedFolder = Me
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overrides Function ToString() As String
      Return Me.Path
    End Function

#End Region

#Region "Protected Methods"

    Protected Friend Overrides Function DebuggerIdentifier() As String
      Try
        If Not String.IsNullOrEmpty(Name) Then
          Return Name
        Else
          Return Path
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Sub SetPath(ByVal lpPath As String)
      mstrPath = lpPath
    End Sub

#End Region

    '#Region "IComparable Implementation"

    '    Public Function CompareTo(ByVal obj As Object) As Integer Implements IComparable.CompareTo

    '      Try
    '        If TypeOf obj Is IFolder Then
    '          Return Name.CompareTo(obj.Name)
    '        Else
    '          Throw New ArgumentException("Object is not a CFolder")
    '        End If
    '      Catch ex As Exception
    '        ApplicationLogging.LogException(ex, String.Format("{0}::CFolder::CompareTo", Me.GetType.Name))
    '      End Try

    '    End Function

    '#End Region

  End Class

End Namespace
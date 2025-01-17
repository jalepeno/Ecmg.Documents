'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities
Imports Documents.Core
Imports Documents.Providers

Namespace Migrations

  ''' <summary>Helper object used to build destination folder paths.</summary>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class PathFactory
    ' Utility class for manipulating folder paths

#Region "Class Variables"

    Private mstrBaseFolderPath As String = String.Empty
    Private mstrDelimiter As String = String.Empty
    Private mblnLeadingDelimiter As Boolean
    Private menuBasePathLocation As ePathLocation
    Private mstrSourceFolderPath As String
    Private menuFilingMode As Core.FilingMode
    Private mblnIncludeDriveInOutPaths As Boolean

#End Region

#Region "Public Enums"

    'Public Enum ePathLocation
    '  Front = 0
    '  Back = 1
    'End Enum

    'Public Enum FolderDelimiterType
    '  ForwardSlash = 0
    '  BackSlash = 1
    'End Enum

#End Region

#Region "Public Properties"

    ''' <summary>The base destination folder path.</summary>
    Public Property BaseFolderPath() As String
      Get
        Return mstrBaseFolderPath
      End Get
      Set(ByVal value As String)
        mstrBaseFolderPath = value
      End Set
    End Property

    ''' <summary>
    ''' Describes whether the path component will be placed in the front of or the back
    ''' of the existing path. The possible values are <strong>Front</strong> and
    ''' <strong>Back</strong>.
    ''' </summary>
    Public Property BasePathLocation() As ePathLocation
      Get
        Return menuBasePathLocation
      End Get
      Set(ByVal value As ePathLocation)
        menuBasePathLocation = value
      End Set
    End Property

    ''' <summary>
    ''' Specifies the character(s) to be used as a folder delimiter in the destination
    ''' repository.
    ''' </summary>
    Public Property Delimiter() As String
      Get
        Return mstrDelimiter
      End Get
      Set(ByVal value As String)
        mstrDelimiter = value
      End Set
    End Property

    ''' <summary>
    ''' Specifies how the document will be filed in the destination. Possible values are
    ''' <strong>UnFiled</strong>, <strong>BaseFolderPathOnly</strong>,
    ''' <strong>DocumentFolderPath</strong>,
    ''' <strong>BaseFolderPathPlusDocumentFolderPath</strong>,
    ''' <strong>DocumentFolderPathPlusBaseFolderPath</strong>.
    ''' </summary>
    Public Property FilingMode() As Core.FilingMode
      Get
        Return menuFilingMode
      End Get
      Set(ByVal value As Core.FilingMode)
        menuFilingMode = value
      End Set
    End Property

    ''' <summary>
    ''' Specifies whether or not the destination folder path will have a leading folder
    ''' delimiter.
    ''' </summary>
    Public Property LeadingDelimiter() As Boolean
      Get
        Return mblnLeadingDelimiter
      End Get
      Set(ByVal value As Boolean)
        mblnLeadingDelimiter = value
      End Set
    End Property

    ''' <summary>The fully qualified path to the source folder.</summary>
    Public Property SourceFolderPath() As String
      Get
        Return mstrSourceFolderPath
      End Get
      Set(ByVal value As String)
        mstrSourceFolderPath = value
      End Set
    End Property

    ''' <summary>
    ''' Allows the drive letter to be included in the destination folder path. Normally
    ''' used when the destination is the file system.
    ''' </summary>
    Public Property IncludeDriveInOutPaths() As Boolean
      Get
        Return mblnIncludeDriveInOutPaths
      End Get
      Set(ByVal value As Boolean)
        mblnIncludeDriveInOutPaths = value
      End Set
    End Property
#End Region

#Region "Constructors"

    ''' <summary>Default Constructor</summary>
    Public Sub New()
      Me.New("")
    End Sub

    ''' <summary>Creates a new PathFactory Object with the specified parameters.</summary>
    Public Sub New(ByVal lpSourceFolderPath As String,
                   Optional ByVal lpBaseFolderPath As String = "",
                   Optional ByVal lpBasePathLocation As ePathLocation = ePathLocation.Front,
                   Optional ByVal lpDelimiter As String = "/",
                   Optional ByVal lpLeadingDelimiter As Boolean = True,
                   Optional ByVal lpFilingMode As Core.FilingMode = Core.FilingMode.UnFiled,
                   Optional ByVal lpIncludeDriveInOutPaths As Boolean = False)

      Try
        SourceFolderPath = lpSourceFolderPath
        BaseFolderPath = lpBaseFolderPath
        BasePathLocation = lpBasePathLocation
        Delimiter = lpDelimiter
        LeadingDelimiter = lpLeadingDelimiter
        FilingMode = lpFilingMode
        IncludeDriveInOutPaths = lpIncludeDriveInOutPaths
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Sub

#End Region

#Region "Public Methods"

    Public Shared Function Create(ByVal lpContentSource As ContentSource,
                                  ByVal lpDocument As Document,
                                  lpFilingMode As FilingMode,
                                  Optional lpLeadingFolderDelimiter As Boolean = True,
                                  Optional lpFolderDelimiter As String = "/",
                                  Optional lpBasePathLocation As ePathLocation = ePathLocation.Front) As PathFactory
      Try
        Dim lobjFolderPaths As IList(Of String) = lpDocument.FolderPaths
        Dim lstrBaseFolderPath As String = String.Empty
        Dim lobjPathFactory As PathFactory = Nothing

        If (lobjFolderPaths Is Nothing) Then
          If (lpLeadingFolderDelimiter = True) Then
            lstrBaseFolderPath = lpFolderDelimiter

          Else
            lstrBaseFolderPath = String.Empty
          End If

          lpFilingMode = FilingMode.UnFiled

        Else

          If (lobjFolderPaths.Count > 0) Then
            lstrBaseFolderPath = lobjFolderPaths.First

          Else

            If (lpLeadingFolderDelimiter = True) Then
              lstrBaseFolderPath = lpFolderDelimiter

            Else
              lstrBaseFolderPath = String.Empty
            End If

            lpFilingMode = FilingMode.UnFiled
          End If

        End If

        Dim lstrOriginalFolderPath As String
        lstrOriginalFolderPath = lstrBaseFolderPath

        If lpContentSource.ProviderName = "File System Provider" Then
          ' This is a file system provider, we need to keep the drive information
          lobjPathFactory = New PathFactory(lstrOriginalFolderPath, lstrBaseFolderPath, lpBasePathLocation, lpFolderDelimiter, False, lpFilingMode, True)

        Else
          ' This is not a file system provider, we need to discard the drive information
          lobjPathFactory = New PathFactory(lstrOriginalFolderPath, lstrBaseFolderPath, lpBasePathLocation, lpFolderDelimiter, lpLeadingFolderDelimiter, lpFilingMode, False)
        End If

        Return lobjPathFactory

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>Creates a folder path using the instance information.</summary>
    Public Function CreateFolderPath() As String

      Try
        Select Case FilingMode
          Case Core.FilingMode.BaseFolderPathOnly
            Return CreateFolderPath("", BaseFolderPath, BasePathLocation, Delimiter, LeadingDelimiter, IncludeDriveInOutPaths)

          Case Core.FilingMode.DocumentFolderPath
            Return CreateFolderPath(SourceFolderPath, "", BasePathLocation, Delimiter, LeadingDelimiter, IncludeDriveInOutPaths)

          Case Core.FilingMode.BaseFolderPathPlusDocumentFolderPath
            Return CreateFolderPath(SourceFolderPath, BaseFolderPath, ePathLocation.Front, Delimiter, LeadingDelimiter, IncludeDriveInOutPaths)

          Case Core.FilingMode.DocumentFolderPathPlusBaseFolderPath
            Return CreateFolderPath(SourceFolderPath, BaseFolderPath, ePathLocation.Back, Delimiter, LeadingDelimiter, IncludeDriveInOutPaths)

          Case Core.FilingMode.UnFiled
            Return ""

          Case Else
            Return ""
        End Select

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Throw New Exception("Unable to create folder path", ex)
      End Try

    End Function

#Region "Shared Methods"

    ''' <summary>Creates a folder path using the specified parameters.</summary>
    ''' <param name="lpSourceDocumentFolderPath">The fully qualified path to the source folder.</param>
    ''' <param name="lpBaseFolderPath">The base destination folder path.</param>
    ''' <param name="lpAddPathLocation">
    ''' Describes whether the path component will be placed in the front of or the back
    ''' of the existing path. The possible values are <strong>Front</strong> and
    ''' <strong>Back</strong>.
    ''' </param>
    ''' <param name="lpDelimiter">
    ''' Specifies the character(s) to be used as a folder delimiter in the destination
    ''' repository.
    ''' </param>
    ''' <param name="lpLeadingDelimiter">
    ''' Specifies whether or not the destination folder path will have a leading folder
    ''' delimiter.
    ''' </param>
    ''' <param name="lpIncludeDriveInPath">
    ''' Allows the drive letter to be included in the destination folder path. Normally
    ''' used when the destination is the file system.
    ''' </param>
    Public Shared Function CreateFolderPath(ByVal lpSourceDocumentFolderPath As String,
                                            Optional ByVal lpBaseFolderPath As String = "",
                                            Optional ByVal lpAddPathLocation As ePathLocation = ePathLocation.Front,
                                            Optional ByVal lpDelimiter As String = "/",
                                            Optional ByVal lpLeadingDelimiter As Boolean = True,
                                            Optional ByVal lpIncludeDriveInPath As Boolean = False) As String

      Dim lstrFolderPath As String = ""
      Dim lstrBasePathCollection As Collections.Specialized.StringCollection = CreateFolderPathCollection(lpBaseFolderPath, lpIncludeDriveInPath)
      Dim lstrSourceDocumentFolderPathCollection As Collections.Specialized.StringCollection = CreateFolderPathCollection(lpSourceDocumentFolderPath, lpIncludeDriveInPath)

      If lpAddPathLocation = ePathLocation.Front Then
        ' We will start from the base folder path
        lstrFolderPath = CreateFolderPath(lstrBasePathCollection, lpDelimiter, lpLeadingDelimiter)
        lstrFolderPath &= CreateFolderPath(lstrSourceDocumentFolderPathCollection, lpDelimiter)
      Else
        ' We will start from the source document folder path
        lstrFolderPath = CreateFolderPath(lstrSourceDocumentFolderPathCollection, lpDelimiter, lpLeadingDelimiter)
        lstrFolderPath &= CreateFolderPath(lstrBasePathCollection, lpDelimiter)
      End If

      ' Eliminate any instances of double delimiters
      lstrFolderPath = lstrFolderPath.Replace(lpDelimiter & lpDelimiter, lpDelimiter)

      Return lstrFolderPath

    End Function

    ''' <summary>
    ''' Creates a folder path using the specified folder collection and the specified
    ''' delimiter parameters.
    ''' </summary>
    ''' <param name="lpPathCollection">The collection of folder names used to construct the complete folder path.</param>
    ''' <param name="lpDelimiter">
    ''' Specifies the character(s) to be used as a folder delimiter in the destination
    ''' repository.
    ''' </param>
    ''' <param name="lpLeadingDelimiter">
    ''' Specifies whether or not the destination folder path will have a leading folder
    ''' delimiter.
    ''' </param>
    Public Shared Function CreateFolderPath(ByVal lpPathCollection As Collections.Specialized.StringCollection, ByVal lpDelimiter As String,
                                            Optional ByVal lpLeadingDelimiter As Boolean = True) As String

      Dim lstrFolderPath As String = ""

      ' If necessary, add the leading delimiter
      If lpLeadingDelimiter = True Then
        lstrFolderPath &= lpDelimiter
      End If

      For Each lstrFolder As String In lpPathCollection
        lstrFolderPath &= lstrFolder & lpDelimiter
      Next

      If lstrFolderPath = lpDelimiter Then
        Return ""
      End If

      lstrFolderPath = lstrFolderPath.Remove(lstrFolderPath.Length - lpDelimiter.Length)

      Return lstrFolderPath

    End Function

    ''' <summary>
    ''' Takes a folder path expressed as a string and splits it into a collection of
    ''' folder strings.
    ''' </summary>
    ''' <returns>A StringCollection containing a set of individual folder names.</returns>
    ''' <param name="lpFolderPath">The folder path to split.</param>
    ''' <param name="lpIncludeDrive">
    ''' Specifies whether or not the drive letter (if applicable) is included in to
    ''' returned collection.
    ''' </param>
    Public Shared Function CreateFolderPathCollection(ByVal lpFolderPath As String,
                                                      Optional ByVal lpIncludeDrive As Boolean = False) As Collections.Specialized.StringCollection

      Dim lobjFolderPathArray As New Collections.Specialized.StringCollection
      Dim lstrForwardFolderPaths As String()
      Dim lstrBackwardFolderPaths As String()
      Dim lstrFolderPaths As String()


      If lpFolderPath Is Nothing Then
        Return New Collections.Specialized.StringCollection
      End If

      If lpFolderPath.Length = 0 Then
        Return New Collections.Specialized.StringCollection
      End If

      Try
        lstrForwardFolderPaths = lpFolderPath.Split("/")
        lstrBackwardFolderPaths = lpFolderPath.Split("\")

        If lstrForwardFolderPaths.Length > lstrBackwardFolderPaths.Length Then
          lstrFolderPaths = lstrForwardFolderPaths
        Else
          lstrFolderPaths = lstrBackwardFolderPaths
        End If

        If lpIncludeDrive = False Then
          ' If we are excluding the drive information then look for the semi-colon and exclude if it is there.
          If lstrFolderPaths(0).EndsWith(":") = False AndAlso lstrFolderPaths(0).Length > 0 Then
            lobjFolderPathArray.Add(lstrFolderPaths(0))
          End If
        Else
          ' If we are not excluding the drive information then we will always add it.
          lobjFolderPathArray.Add(lstrFolderPaths(0))
        End If

        If lstrFolderPaths.Length > 1 Then
          For lintFolderCounter As Integer = 1 To lstrFolderPaths.Length - 1
            lobjFolderPathArray.Add(lstrFolderPaths(lintFolderCounter))
          Next
        End If

        Return lobjFolderPathArray

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return New Collections.Specialized.StringCollection
      End Try

    End Function

#End Region

#End Region

#Region "Private Methods"

    Protected Friend Overridable Function DebuggerIdentifier() As String
      Try
        Return CreateFolderPath()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
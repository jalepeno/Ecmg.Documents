'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports Documents.Core
Imports Documents.SerializationUtilities
Imports Documents.Utilities
Imports Ionic.Zip

#End Region

Namespace Comparison

  ''' <summary>
  ''' Used to compare to document objects.
  ''' </summary>
  ''' <remarks></remarks>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Partial Public Class DocumentComparison
    Inherits CtsObject
    Implements IDisposable
    Implements INotifyPropertyChanged

#Region "Public Constants"

    ''' <summary>
    ''' Constant value representing the file extension to save XML serialized instances
    ''' to.
    ''' </summary>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public Const DOCUMENT_COMPARISON_FILE_EXTENSION As String = "cmp"

    ''' <summary>
    ''' Constant value representing the file extension to save XML serialized instances
    ''' to.
    ''' </summary>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Public Const DOCUMENT_COMPARISON_ARCHIVE_FILE_EXTENSION As String = "cfa"

#End Region

#Region "Public Events"

    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

#End Region

#Region "Class Variables"

    Private mstrId As String = String.Empty
    Private mobjDocumentX As Document = Nothing
    Private mobjDocumentY As Document = Nothing

#End Region

#Region "Public Properties"

    Public Property ID As String
      Get
        Try
          If String.IsNullOrEmpty(mstrId) Then
            mstrId = GenerateId()
          End If
          Return mstrId
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        mstrId = value
        OnPropertyChanged("ID")
      End Set
    End Property

    Public Property DocumentX As Document
      Get
        Return mobjDocumentX
      End Get
      Set(value As Document)
        Try
          mobjDocumentX = value
          mstrDocumentXID = value.ID
          OnPropertyChanged("DocumentX")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property DocumentY As Document
      Get
        Return mobjDocumentY
      End Get
      Set(value As Document)
        Try
          mobjDocumentY = value
          mstrDocumentYID = value.ID
          OnPropertyChanged("DocumentY")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property AllProperties As New ComparisonProperties

    Public ReadOnly Property PropertiesInBothDocuments As ComparisonProperties
      Get
        Return GetProperties(ComparisonProperty.MatchStatusEnum.ExistsInBothDocuments)
      End Get
    End Property

    Public ReadOnly Property PropertiesInXOnly As ComparisonProperties
      Get
        Return GetProperties(ComparisonProperty.MatchStatusEnum.DocumentXOnly)
      End Get
    End Property

    Public ReadOnly Property PropertiesInYOnly As ComparisonProperties
      Get
        Return GetProperties(ComparisonProperty.MatchStatusEnum.DocumentYOnly)
      End Get
    End Property

    Public ReadOnly Property EqualProperties As ComparisonProperties
      Get
        Return GetEqualProperties()
      End Get
    End Property

    Public ReadOnly Property UnequalProperties As ComparisonProperties
      Get
        Return GetUnequalProperties()
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(lpDocumentXPath As String, lpDocumentYPath As String)
      Try
        DocumentX = New Document(lpDocumentXPath)
        DocumentY = New Document(lpDocumentYPath)
        Compare()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(x As Document, y As Document)
      Try
        DocumentX = x
        DocumentY = y
        Compare()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(lpFilePath As String)
      Try
        Dim lstrErrorMessage As String = String.Empty
        Dim lobjDocumentComparison As DocumentComparison = Deserialize(lpFilePath, lstrErrorMessage)

        If lobjDocumentComparison Is Nothing Then
          ' Check the error message
          If Not String.IsNullOrEmpty(lstrErrorMessage) Then
            Throw New Exception(lstrErrorMessage)
          End If
        End If

        Helper.AssignObjectProperties(lobjDocumentComparison, Me)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Executes the comparison logic for the two documents.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Compare()
      Try

        Dim lobjComparisonProperty As ComparisonProperty = Nothing

        ' Make sure we have both documents to compare
        If DocumentX Is Nothing AndAlso DocumentY Is Nothing Then
          Throw New InvalidOperationException("Unable to compare documents, neither document has been specified.")
        End If

        If DocumentX Is Nothing And DocumentY IsNot Nothing Then
          Throw New InvalidOperationException("Unable to compare documents, DocumentX has not been specified.")
        End If

        If DocumentX IsNot Nothing And DocumentY Is Nothing Then
          Throw New InvalidOperationException("Unable to compare documents, DocumentY has not been specified.")
        End If

        AllProperties.Clear()

        ' Add the document level properties for document x
        For Each lobjProperty As IProperty In DocumentX.Properties
          Me.AllProperties.Add(New ComparisonProperty(lobjProperty, PropertyScope.DocumentProperty))
        Next

        ' Add the document level properties for document y
        For Each lobjProperty As IProperty In DocumentY.Properties
          ' Clear the comparison property
          lobjComparisonProperty = Nothing
          lobjComparisonProperty = AllProperties.Item(lobjProperty.Name, PropertyScope.DocumentProperty, String.Empty)
          If lobjComparisonProperty IsNot Nothing Then
            lobjComparisonProperty.PropertyY = lobjProperty
          Else
            lobjComparisonProperty = New ComparisonProperty
            With lobjComparisonProperty
              .PropertyY = lobjProperty
              .Scope = PropertyScope.DocumentProperty
            End With
            AllProperties.Add(lobjComparisonProperty)
          End If
        Next

        ' Add the version properties for document x
        For Each lobjVersion As Version In DocumentX.Versions
          For Each lobjProperty As IProperty In lobjVersion.Properties
            AllProperties.Add(New ComparisonProperty(lobjProperty, PropertyScope.VersionProperty, lobjVersion.ID))
          Next
        Next

        ' Add the version properties for document x
        For Each lobjVersion As Version In DocumentY.Versions
          For Each lobjProperty As IProperty In lobjVersion.Properties
            ' Clear the comparison property
            lobjComparisonProperty = Nothing
            lobjComparisonProperty = AllProperties.Item(lobjProperty.Name, PropertyScope.VersionProperty, lobjVersion.ID)
            If lobjComparisonProperty IsNot Nothing Then
              lobjComparisonProperty.PropertyY = lobjProperty
            Else
              lobjComparisonProperty = New ComparisonProperty
              With lobjComparisonProperty
                .PropertyY = lobjProperty
                .Scope = PropertyScope.VersionProperty
                .VersionId = lobjVersion.ID
              End With
              AllProperties.Add(lobjComparisonProperty)
            End If
          Next
        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shared Function Compare(lpDocumentXPath As String, lpDocumentYPath As String) As DocumentComparison
      Try
        Dim lobjDocumentComparison As New DocumentComparison(lpDocumentXPath, lpDocumentYPath)

        Return lobjDocumentComparison

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function Compare(x As Document, y As Document) As DocumentComparison
      Try
        Dim lobjDocumentComparison As New DocumentComparison(x, y)
        'lobjDocumentComparison.Compare()

        Return lobjDocumentComparison

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub Save(ByRef lpFilePath As String)
      Try
        Save(lpFilePath, True)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Save(ByRef lpFilePath As String, lpPassword As String)
      Try
        Save(lpFilePath, True, lpPassword)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Save(ByRef lpFilePath As String, lpIncludeDocuments As Boolean)
      Try
        Save(lpFilePath, lpIncludeDocuments, String.Empty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Save(ByRef lpFilePath As String, lpIncludeDocuments As Boolean, lpPassword As String)
      Try

        Dim lobjArchiveMemoryStream As MemoryStream = ToArchiveStream(lpPassword)
        Dim lobjArchiveFileStream As New FileStream(lpFilePath, FileMode.Create, FileAccess.Write)

        lobjArchiveMemoryStream.WriteTo(lobjArchiveFileStream)

        lobjArchiveFileStream.Close()

        lobjArchiveMemoryStream.Dispose()
        lobjArchiveFileStream.Dispose()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function FromArchive(ByVal lpFilePath As String) As DocumentComparison
      Try
        Dim lobjReturnObject As DocumentComparison = Nothing
        Dim lobjZipStream As MemoryStream = Nothing
        Dim lobjDocumentX As New Document
        Dim lobjDocumentY As New Document

        Using lobjZipFile As New ZipFile(lpFilePath)
          'Dim lstrFileName As String = IO.Path.GetFileName(lpFilePath).Replace(DefaultArchiveFileExtension, DefaultFileExtension)
          For Each lobjEntry As ZipEntry In lobjZipFile.Entries
            lobjZipStream = New MemoryStream
            lobjEntry.Extract(lobjZipStream)
            If String.Equals(IO.Path.GetExtension(lobjEntry.FileName).TrimStart("."), DefaultFileExtension) Then
              lobjReturnObject = Serializer.Deserialize.FromZippedStream(lobjZipStream, lobjZipFile, Me.GetType)
            Else
              If lobjEntry.FileName.StartsWith(lobjReturnObject.mstrDocumentXID) Then
                lobjReturnObject.DocumentX = New Document(lobjZipStream)
              End If
              If lobjEntry.FileName.StartsWith(lobjReturnObject.mstrDocumentYID) Then
                lobjReturnObject.DocumentY = New Document(lobjZipStream)
              End If
            End If
          Next
        End Using
        Return lobjReturnObject
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function FromArchiveStream(lpStream As IO.Stream) As DocumentComparison
      Try

        Using lobjZipFile As ZipFile = ZipFile.Read(lpStream)
          Return Serializer.Deserialize.FromZippedStream(lpStream, lobjZipFile, Me.GetType)
        End Using

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToArchiveStream(ByVal password As String) As IO.Stream
      Try
        Dim lobjOutputStream As New MemoryStream
        Dim lobjZipFile As New ZipFile
        'Dim lobjDocumentStreams As New List(Of IO.Stream)

        With lobjZipFile

          '.AddEntry(String.Format("{0}.{1}", Me.ID, Me.DefaultFileExtension), Me.ToStream)
          .AddEntry(String.Format("Comparison {0}.{1}", Me.ID, Me.DefaultFileExtension), Me.ToStream)

          ' Get the documents
          If DocumentX IsNot Nothing Then
            .AddEntry(String.Format("{0}.{1}", DocumentX.ID, DocumentX.DefaultArchiveFileExtension), DocumentX.ToArchiveStream)
            'lobjDocumentStreams.Add(DocumentX.ToArchiveStream(password))
          End If

          If DocumentY IsNot Nothing Then
            .AddEntry(String.Format("{0}.{1}", DocumentY.ID, DocumentY.DefaultArchiveFileExtension), DocumentY.ToArchiveStream)
            'lobjDocumentStreams.Add(DocumentY.ToArchiveStream(password))
          End If

          If Not String.IsNullOrEmpty(password) Then
            .Password = password
          End If

          .Save(lobjOutputStream)

        End With

        If lobjOutputStream.CanSeek Then
          lobjOutputStream.Position = 0
        End If

        lobjZipFile.Dispose()

        lobjZipFile = Nothing

        Return lobjOutputStream

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overrides Function ToString() As String
      Try
        Return DebuggerIdentifier()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToStream() As IO.MemoryStream
      Try

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        Return Serializer.Serialize.ToStream(Me)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Properties"

    Private ReadOnly Property IsDisposed() As Boolean
      Get
        Return disposedValue
      End Get
    End Property

#End Region

#Region "Private Methods"

    Private Function GetProperties(lpMatchStatus As ComparisonProperty.MatchStatusEnum) As ComparisonProperties
      Try
        Dim lobjReturnProperties As New ComparisonProperties
        Dim list As Object = From lobjProperty In AllProperties Where
               lobjProperty.MatchStatus = lpMatchStatus Select lobjProperty
        lobjReturnProperties.Add(list)
        Return lobjReturnProperties
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function GetEqualProperties() As ComparisonProperties
      Try
        Dim lobjReturnProperties As New ComparisonProperties
        Dim list As Object = From lobjProperty In AllProperties Where
               lobjProperty.MatchStatus = ComparisonProperty.MatchStatusEnum.ExistsInBothDocuments _
               And lobjProperty.IsEqual = True Select lobjProperty
        lobjReturnProperties.Add(list)
        Return lobjReturnProperties
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function GetUnequalProperties() As ComparisonProperties
      Try
        Dim lobjReturnProperties As New ComparisonProperties
        Dim list As Object = From lobjProperty In AllProperties Where
               lobjProperty.MatchStatus = ComparisonProperty.MatchStatusEnum.ExistsInBothDocuments _
               And lobjProperty.IsEqual = False Select lobjProperty
        lobjReturnProperties.Add(list)
        Return lobjReturnProperties
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Overridable Function DebuggerIdentifier() As String
      Try
        Dim lobjReturnBuilder As New StringBuilder

        lobjReturnBuilder.AppendFormat("{0} Matching Properties: {1} Equal / {2} Unequal", PropertiesInBothDocuments.Count, EqualProperties.Count, UnequalProperties.Count)

        Return lobjReturnBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function GenerateId() As String
      Try
        Dim lobjStringBuilder As New StringBuilder

        If DocumentX IsNot Nothing Then
          lobjStringBuilder.Append(DocumentX.ID)
          If DocumentY IsNot Nothing Then
            lobjStringBuilder.AppendFormat("^{0}", DocumentY.ID)
          End If
        Else
          If DocumentY IsNot Nothing Then
            lobjStringBuilder.Append(DocumentY.ID)
          Else
            lobjStringBuilder.Append("NotInitialized")
          End If
        End If

        Return lobjStringBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Sub OnPropertyChanged(ByVal lpPropertyName As String)
      Try
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(lpPropertyName))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
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
          ' DISPOSETODO: dispose managed state (managed objects).
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
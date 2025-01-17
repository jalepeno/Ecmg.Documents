'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Data
Imports Documents.SerializationUtilities
Imports Documents.Utilities

#End Region

Namespace Core
  ''' <summary>
  ''' Contains a set of search results returned from a search
  ''' </summary>
  ''' <remarks>If the search operation resulted in an exception the HasException property will 
  ''' have a value of True and the Exception property will contain a 
  ''' reference to the exception thrown by the search.</remarks>
  Public Class SearchResultSet
    'Inherits Search.SearchResultSet
    Implements IDisposable
    Implements ISerialize
    Implements ITableable

#Region "Class Constants"

    ''' <summary>
    ''' Constant value representing the 
    ''' file extension to save XML 
    ''' serialized instances to.
    ''' </summary>
    Public Const SEARCH_RESULT_SET_FILE_EXTENSION As String = "srs"

#End Region

#Region "Class Variables"

    Private mobjResults As New SearchResults
    Private mobjException As Exception
    Private mblnHasException As Boolean
    Private mobjDataTable As DataTable

#End Region

#Region "Public Properties"

    Public ReadOnly Property HasException() As Boolean
      Get

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        Return mblnHasException

      End Get
    End Property

    Public ReadOnly Property Results() As SearchResults
      Get

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        Return mobjResults

      End Get
    End Property

    Public ReadOnly Property Exception() As Exception
      Get

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        Return mobjException

      End Get
    End Property

    Public ReadOnly Property Count() As Integer
      Get

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        Return Results.Count

      End Get
    End Property

#End Region

#Region "Private Properties"

    Private ReadOnly Property IsDisposed() As Boolean
      Get
        Return disposedValue
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpException As Exception)
      Me.New(New SearchResults, lpException)
    End Sub

    Public Sub New(ByVal lpResults As SearchResults)
      Me.New(lpResults, Nothing)
    End Sub

    Public Sub New(ByVal lpResults As SearchResults, ByVal lpException As Exception)
      If lpException IsNot Nothing Then
        mobjException = lpException
        mblnHasException = True
      Else
        mblnHasException = False
      End If

      mobjResults = lpResults

    End Sub

#End Region

#Region "ISerialize Implementation"

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
        Return SEARCH_RESULT_SET_FILE_EXTENSION
      End Get
    End Property

    ''' <summary>
    ''' Instantiate from an XML file.
    ''' </summary>
    Public Function Deserialize(ByVal lpFilePath As String, Optional ByRef lpErrorMessage As String = "") As Object Implements ISerialize.Deserialize
      Try

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        Return Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        lpErrorMessage = Helper.FormatCallStack(ex)
        Return Nothing
      End Try
    End Function

    Public Function DeSerialize(ByVal lpXML As System.Xml.XmlDocument) As Object Implements ISerialize.Deserialize
      Try

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        Return Serializer.Deserialize.XmlString(lpXML.OuterXml, Me.GetType)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Helper.DumpException(ex)
        Return Nothing
      End Try
    End Function

    ''' <summary>
    ''' Saves a representation of the object in an XML file.
    ''' </summary>
    ''' <param name="lpFilePath">The fully qualified path to save the object to.</param>
    Public Sub Serialize(ByVal lpFilePath As String) Implements ISerialize.Serialize
      Try

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        Serialize(lpFilePath, "")

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Saves a representation of the object in an XML file.
    ''' </summary>
    ''' <param name="lpFilePath">The fully qualified path to save the object to.</param>
    ''' <param name="lpFileExtension"></param>
    ''' <remarks></remarks>
    Public Sub Serialize(ByRef lpFilePath As String, ByVal lpFileExtension As String) Implements ISerialize.Serialize
      ' lpFileExtension ignored in this implementation

      If IsDisposed Then
        Throw New ObjectDisposedException(Me.GetType.ToString)
      End If

      Serializer.Serialize.XmlFile(Me, lpFilePath)

    End Sub

    Public Sub Serialize(ByVal lpFilePath As String, ByVal lpWriteProcessingInstruction As Boolean, ByVal lpStyleSheetPath As String) Implements ISerialize.Serialize

      If IsDisposed Then
        Throw New ObjectDisposedException(Me.GetType.ToString)
      End If

      If lpWriteProcessingInstruction = True Then
        Serializer.Serialize.XmlFile(Me, lpFilePath, , , True, lpStyleSheetPath)
      Else
        Serializer.Serialize.XmlFile(Me, lpFilePath)
      End If

    End Sub

    Public Function Serialize() As System.Xml.XmlDocument Implements ISerialize.Serialize

      If IsDisposed Then
        Throw New ObjectDisposedException(Me.GetType.ToString)
      End If

      Return Serializer.Serialize.Xml(Me)

    End Function

    ''' <summary>
    ''' Returns an XML serialized rendition of the SearchResultSet
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ToXmlString() As String Implements ISerialize.ToXmlString

      If IsDisposed Then
        Throw New ObjectDisposedException(Me.GetType.ToString)
      End If

      Return Serializer.Serialize.XmlString(Me)

    End Function

#End Region

#Region "ITableable Implementation"

    Public Function ToDataTable() As DataTable Implements ITableable.ToDataTable
      Try

        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If

        If mobjDataTable IsNot Nothing Then
          Return mobjDataTable
        End If

        ' If there were no results then we will return a null reference
        If Results.Count = 0 Then
          Return Nothing
        End If

        ' Build the data table

        mobjDataTable = New DataTable("SearchResultSet")

        Dim lobjDataRow As DataRow

        With mobjDataTable

          ' Add our columns
          For Each lobjDataItem As Data.DataItem In Me.Results(0).Values
            Select Case lobjDataItem.Type
              Case PropertyType.ecmString, PropertyType.ecmGuid
                .Columns.Add(lobjDataItem.Name, GetType(String))

              Case PropertyType.ecmDate
                .Columns.Add(lobjDataItem.Name, GetType(DateTime))

              Case PropertyType.ecmLong
                .Columns.Add(lobjDataItem.Name, GetType(Integer))

              Case PropertyType.ecmDouble
                .Columns.Add(lobjDataItem.Name, GetType(Double))

              Case PropertyType.ecmBoolean
                .Columns.Add(lobjDataItem.Name, GetType(Boolean))

                'Case PropertyType.ecmGuid
                '  .Columns.Add(lobjDataItem.Name, GetType(Guid))

              Case PropertyType.ecmObject
                .Columns.Add(lobjDataItem.Name, GetType(Object))

              Case PropertyType.ecmBinary
                .Columns.Add(lobjDataItem.Name, GetType(Object))

            End Select
          Next

          ' Add the rows
          For Each lobjResult As SearchResult In Me.Results
            lobjDataRow = mobjDataTable.NewRow
            For Each lobjDataItem As Data.DataItem In lobjResult.Values

              If (lobjDataItem.Value IsNot Nothing) Then
                lobjDataRow(lobjDataItem.Name) = lobjDataItem.Value
              Else
                lobjDataRow(lobjDataItem.Name) = DBNull.Value
              End If

            Next

            .Rows.Add(lobjDataRow)

          Next

        End With

        Return mobjDataTable

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region " IDisposable Support "

    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
      Try
        If Not Me.disposedValue Then
          If disposing Then
            ' DISPOSETODO: free other state (managed objects).
            mobjDataTable.Dispose()
          End If

          ' DISPOSETODO: free your own state (unmanaged objects).
          ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
      Catch ex As Exception
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

  End Class

End Namespace
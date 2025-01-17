' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ContentPropertyAccessor.vb
'  Description :  [type_description_here]
'  Created     :  7/20/2012 9:39:50 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.IO
Imports Documents.Utilities

#End Region

Namespace Core

  Public MustInherit Class ContentPropertyAccessor
    Implements IContentPropertyAccessor

#Region "Class Variables"

    Private mstrFilePath As String
    Private mstrIdentifier As String
    Protected WithEvents mobjProperties As New ClassificationProperties
    Private mstrTempFilePath As String = String.Empty

#End Region

#Region "Public Properties"

    Public ReadOnly Property FilePath As String
      Get
        Return mstrFilePath
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(lpFilePath As String)
      Try
        mstrFilePath = lpFilePath
        mstrIdentifier = lpFilePath
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(lpContent As Content)
      Try
        mstrIdentifier = lpContent.FileName
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "IContentPropertyAccessor Implementation"

    Public MustOverride Sub Close() Implements Core.IContentPropertyAccessor.Close

    Public MustOverride ReadOnly Property IsReadOnly As Boolean Implements Core.IContentPropertyAccessor.IsReadOnly

    Public Sub Open(lpStream As Stream, lpFileName As String)
      Try
        mstrTempFilePath = Helper.CleanPath(String.Format("{0}\{1}", FileHelper.Instance.TempPath, lpFileName))
        Dim lobjFileStream As New FileStream(mstrFilePath, FileMode.Create)
        If TypeOf lpStream Is MemoryStream Then
          CType(lpStream, MemoryStream).WriteTo(lobjFileStream)
          lobjFileStream.Close()
          Open(lpFileName)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public MustOverride Sub Open(lpFilePath As String) Implements Core.IContentPropertyAccessor.Open

    Public MustOverride Sub Save() Implements Core.IContentPropertyAccessor.Save

    Public MustOverride Sub SyncronizeContentProperties(lpSource As IMetaHolder) Implements IContentPropertyAccessor.SyncronizeContentProperties

    Public Property Identifier As String Implements Core.IMetaHolder.Identifier
      Get
        Return mstrIdentifier
      End Get
      Set(value As String)
        mstrIdentifier = value
      End Set
    End Property

    Public Property Metadata As Core.IProperties Implements Core.IMetaHolder.Metadata
      Get
        Return mobjProperties
      End Get
      Set(value As Core.IProperties)
        mobjProperties = value
      End Set
    End Property

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
      If Not Me.disposedValue Then
        If disposing Then
          ' DISPOSETODO: dispose managed state (managed objects).
          Close()
          If Not String.IsNullOrEmpty(mstrTempFilePath) AndAlso File.Exists(mstrTempFilePath) Then
            Try
              File.Delete(mstrTempFilePath)
            Catch ex As Exception
              ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            End Try
          End If
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

#End Region

  End Class

End Namespace
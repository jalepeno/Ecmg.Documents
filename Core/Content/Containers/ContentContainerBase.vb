' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ContentContainerBase.vb
'  Description :  [type_description_here]
'  Created     :  12/14/2011 8:13:00 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Core

  Public MustInherit Class ContentContainerBase
    Implements IContentContainer
    Implements IDisposable

#Region "Class Variables"

    Private mstrFileName As String = String.Empty
    Private mstrMimeType As String = String.Empty
    Private mobjFileContent As Object = Nothing

#End Region

#Region "Public Properties"

    Public Property FileName As String Implements IContentContainer.FileName
      Get
        Return mstrFileName
      End Get
      Set(ByVal value As String)
        mstrFileName = value
      End Set
    End Property

    Public Property MimeType As String Implements IContentContainer.MimeType
      Get
        Return mstrMimeType
      End Get
      Set(ByVal value As String)
        mstrMimeType = value
      End Set
    End Property

    Public Property FileContent As Object Implements IContentContainer.FileContent
      Get
        Return mobjFileContent
      End Get
      Set(ByVal value As Object)
        mobjFileContent = value
      End Set
    End Property

    Public MustOverride ReadOnly Property CanRead As Boolean Implements IContentContainer.CanRead

#End Region

#Region "Constructors"

    Protected Sub New()

    End Sub

    Protected Sub New(ByVal lpFileName As String, ByVal lpMimeType As String, ByVal lpFileContent As Object)
      Try
        FileName = lpFileName
        MimeType = lpMimeType
        FileContent = lpFileContent
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Sub New(ByVal lpContent As Content)
      Try
        FileName = lpContent.FileName
        MimeType = lpContent.MIMEType
        If lpContent.CanRead Then
          FileContent = lpContent.ToMemoryStream
        End If
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
          If TypeOf FileContent Is IO.Stream Then
            CType(FileContent, IO.Stream).Close()
            CType(FileContent, IO.Stream).Dispose()
          End If
          mobjFileContent = Nothing
          mstrMimeType = String.Empty
          mstrFileName = String.Empty
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
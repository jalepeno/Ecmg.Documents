'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports System.Text
Imports Documents.Utilities

Namespace Core

  Public Class Contents
    Inherits CCollection(Of Content)

#Region "Class Variables"

    Private mobjVersion As Version

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets the Version to which this Content collection belongs
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Version() As Version
      Get
        Return mobjVersion
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal version As Version)
      Try
        mobjVersion = version
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Overloaded Methods"

    ' <Added by: Ernie at: 9/22/2011-8:01:46 AM on machine: ERNIE-M4400>
    ' We need to make sure that the Version parent is always set.
    Public Overloads Sub Add(ByVal lpContent As Content)
      Try
        lpContent.SetVersion(Me.Version)
        MyBase.Add(lpContent)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub
    ' </Added by: Ernie at: 9/22/2011-8:01:46 AM on machine: ERNIE-M4400>

    Public Overloads Sub Add(ByVal lpContentPath As String)
      Dim lobjContent As New Content(lpContentPath, Me.Version)
      lobjContent.SetVersion(Me.Version)
      MyBase.Add(lobjContent)
    End Sub

    Public Overloads Sub Add(ByVal lpContentPath As String, ByVal lpAllowZeroLengthContent As Boolean)
      Dim lobjContent As New Content(lpContentPath, Me.Version, lpAllowZeroLengthContent)
      lobjContent.SetVersion(Me.Version)
      MyBase.Add(lobjContent)
    End Sub

    Public Overloads Sub Add(ByVal lpContentPath As String, ByVal lpStorageType As Content.StorageTypeEnum)
      Dim lobjContent As New Content(lpContentPath, lpStorageType, Me.Version)
      lobjContent.SetVersion(Me.Version)
      MyBase.Add(lobjContent)
    End Sub

    Public Overloads Sub Add(ByVal lpContentPath As String,
                             ByVal lpStorageType As Content.StorageTypeEnum,
                             ByVal lpAllowZeroLengthContent As Boolean)
      Dim lobjContent As New Content(lpContentPath, lpStorageType, Me.Version, lpAllowZeroLengthContent)
      lobjContent.SetVersion(Me.Version)
      MyBase.Add(lobjContent)
    End Sub

    Public Overloads Sub Add(ByVal lpByteArray As Byte(),
                             ByVal lpContentFileName As String,
                             ByVal lpStorageType As Content.StorageTypeEnum)
      Try
        Dim lobjContent As New Content(lpByteArray, lpContentFileName, lpStorageType)
        lobjContent.SetVersion(Me.Version)
        MyBase.Add(lobjContent)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Sub Add(ByVal lpByteArray As Byte(),
                         ByVal lpContentFileName As String,
                         ByVal lpStorageType As Content.StorageTypeEnum,
                         ByVal lpAllowZeroLengthContent As Boolean)
      Try
        Dim lobjContent As New Content(lpByteArray, lpContentFileName, lpStorageType, False, lpAllowZeroLengthContent)
        lobjContent.SetVersion(Me.Version)
        MyBase.Add(lobjContent)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Sub Add(ByVal lpStream As NamedStream)
      Try
        Add(lpStream.Stream, lpStream.FileName, Content.StorageTypeEnum.Reference)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Sub Add(ByVal lpStream As IO.Stream,
                                ByVal lpContentFileName As String,
                                ByVal lpStorageType As Content.StorageTypeEnum)
      Try
        Dim lobjContent As New Content(lpStream, lpContentFileName, lpStorageType)
        lobjContent.SetVersion(Me.Version)
        MyBase.Add(lobjContent)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Sub Add(ByVal lpStream As IO.Stream,
                             ByVal lpContentFileName As String,
                             ByVal lpStorageType As Content.StorageTypeEnum,
                             ByVal lpAllowZeroLengthContent As Boolean)
      Try
        Dim lobjContent As New Content(lpStream, lpContentFileName, lpStorageType, False, lpAllowZeroLengthContent)
        lobjContent.SetVersion(Me.Version)
        MyBase.Add(lobjContent)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Default Overrides Property Item(ByVal index As Integer) As Content
      Get
        Try
          If ValidateIndex(index) Then
            Return MyBase.Item(index)
          Else
            Return Nothing
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As Content)
        If ValidateIndex(index) Then
          MyBase.Item(index) = value
        End If
      End Set
    End Property

    Private Function ValidateIndex(ByVal index As Integer) As Boolean
      Try
        If index > -1 Then
          If Me.Count - 1 < index Then
            Dim lobjExceptionMessageBuilder As New StringBuilder
            lobjExceptionMessageBuilder.AppendFormat("The requested content element index '{0}' is out of the range of available elements.  ", index)
            Select Case Me.Count
              Case 1
                lobjExceptionMessageBuilder.Append("There is only 1 element.  Please treat this index as a zero based array.")
              Case Else
                lobjExceptionMessageBuilder.Append("There are only {1} elements.  Please treat this index as a zero based array.",
                                          Me.Count - 1)
            End Select
            Throw New Exceptions.ItemDoesNotExistException(index, lobjExceptionMessageBuilder.ToString)
          End If
          Return True
        Else
          Return False
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
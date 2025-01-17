'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports System.Text
Imports Documents.SerializationUtilities
Imports Documents.Utilities

#End Region

Namespace Core

  ''' <summary>Provides an abstract desciption of a repository.</summary>
  ''' <remarks>Serialized instances should use the file extension 
  ''' .rif (Repository Information File) or .rfa (Repository File Archive)</remarks>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class Repository
    Implements INotifyPropertyChanged

#Region "Class Constants"

    Public Const REPOSITORY_INFORMATION_FILE_EXTENSION As String = "rif"
    Public Const REPOSITORY_INFORMATION_FILE_ARCHIVE_EXTENSION As String = "rfa"

#End Region

#Region "Class Variables"

    Private mstrName As String = String.Empty

    Private mobjDocumentClasses As DocumentClasses
    Private mobjProperties As ClassificationProperties
    'Private mstrExportPath As String
    'Private mstrImportPath As String
    Private mintMaxRifClassCount As Integer = 10
    Private mobjSerializationSourceType As SourceType = SourceType.File
    Private mobjRepositoryStream As IO.Stream

#End Region

#Region "Public Events"

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' The name of the repository
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Name() As String
      Get
        Return mstrName
      End Get
      Set(ByVal value As String)
        Try
          If String.Compare(mstrName, value) <> 0 Then
            mstrName = value
            NotifyPropertyChanged("Name")
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' The collection of document classes defined in the repository
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DocumentClasses() As DocumentClasses
      Get
        Try
          If mobjDocumentClasses Is Nothing Then
            mobjDocumentClasses = New DocumentClasses
          End If
          Return mobjDocumentClasses
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As DocumentClasses)
        mobjDocumentClasses = value
      End Set
    End Property

    ''' <summary>
    ''' The complete set of properties defined for the repository
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Properties() As ClassificationProperties
      Get
        Try
          If mobjProperties Is Nothing Then
            mobjProperties = New ClassificationProperties
          End If
          Return mobjProperties
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As ClassificationProperties)
        mobjProperties = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      Try
        Name = String.Empty
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Determine if a property exists in the document.
    ''' </summary>
    Public Function PropertyExists(ByVal lpName As String) As Boolean
      Try
        Return Properties.PropertyExists(lpName)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overrides Function ToString() As String
      Try
        Return Me.DebuggerIdentifier
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Protected Methods"

    Protected Friend Overridable Function DebuggerIdentifier() As String
      Dim lobjIdentifierBuilder As New StringBuilder
      Try

        lobjIdentifierBuilder.AppendFormat("{0}: ", Me.Name)

        Select Case Me.DocumentClasses.Count
          Case 0
            lobjIdentifierBuilder.Append("No Document Classes: ")
          Case 1
            lobjIdentifierBuilder.Append("1 Document Class: ")
          Case Else
            lobjIdentifierBuilder.AppendFormat("{0} Document Classes: ", Me.DocumentClasses.Count)
        End Select

        Select Case Me.Properties.Count
          Case 0
            lobjIdentifierBuilder.Append("No Properties")
          Case 1
            lobjIdentifierBuilder.Append("1 Property")
          Case Else
            lobjIdentifierBuilder.AppendFormat("{0} Properties", Me.Properties.Count)
        End Select

        Return lobjIdentifierBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return lobjIdentifierBuilder.ToString
      End Try
    End Function

#End Region

#Region "Private Methods"

    Private Sub NotifyPropertyChanged(ByVal lpInfo As String)
      Try
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(lpInfo))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace

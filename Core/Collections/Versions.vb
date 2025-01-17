'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Core

  Public Class Versions
    Inherits CCollection(Of Version)

#Region "Class Variables"

    Private mobjDocument As Document

#End Region

#Region "Public Properties"

    Public ReadOnly Property Document() As Document
      Get
        Return mobjDocument
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal document As Document)
      Try
        mobjDocument = document
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Overloads Sub Add(ByVal lpVersion As Version)
      Try
        lpVersion.SetDocument(Me.Document)
        If Helper.CallStackContainsMethodName("Deserialize") = False Then
          lpVersion.ID = Me.Count
        End If
        MyBase.Add(lpVersion)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Checks to see if an item exists with the specified index number.
    ''' </summary>
    ''' <param name="lpId">The index number to check for.</param>
    ''' <returns>True if the index number is valid, otherwise False.</returns>
    ''' <remarks></remarks>
    Public Function IdExists(ByVal lpId As Integer) As Boolean
      Try
        For Each lobjVersion As Version In Me
          If lobjVersion.ID = lpId Then
            Return True
          End If
        Next

        Return False

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

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

#Region "Item Overloads"

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Default Shadows Property Item(ByVal ID As String) As Version
      Get
        Try
          Dim lobjVersion As Version
          For lintCounter As Integer = 0 To MyBase.Count - 1
            lobjVersion = CType(MyBase.Item(lintCounter), Version)
            If lobjVersion.ID = ID Then
              Return lobjVersion
            End If
          Next
          If ID = "0" Then
            ' We did not get it by 0 try 1
            Return Item(1)
          End If
          Throw New Exception("There is no Item by the ID '" & ID & "'.")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("Versions::Get_Item('{0}')", ID))
          ' Re-throw the exception to the caller
          Throw
        End Try
        'Throw New InvalidArgumentException
      End Get
      Set(ByVal value As Version)
        Try
          Dim lobjVersion As Version
          For lintCounter As Integer = 0 To MyBase.Count - 1
            lobjVersion = CType(MyBase.Item(lintCounter), Version)
            If lobjVersion.ID = ID Then
              MyBase.Item(lintCounter) = value
            End If
          Next
          Throw New Exception("There is no Item by the ID '" & ID & "'.")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("Versions::Set_Item('{0}')", ID))
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Default Shadows Property Item(ByVal Index As Integer) As Version
      Get
        'ApplicationLogging.WriteLogEntry("Enter Versions::Item(Index as Integer)", TraceEventType.Verbose)
        Try
          If Index = -1 Then
            ' This is normally a request for all versions, but in this context we can only return one
            ' Return the first version
            If Count > 0 Then
              Return MyBase.Item(0)
            Else
              ' There are no versions
              Return Nothing
            End If
          End If
          Return MyBase.Item(Index)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("Versions::Get_Item('{0}')", Index))
          ' Re-throw the exception to the caller
          Throw
        Finally
          'ApplicationLogging.WriteLogEntry("Exit Versions::Item(Index as Integer)", TraceEventType.Verbose)
        End Try
      End Get
      Set(ByVal value As Version)
        MyBase.Item(Index) = value
      End Set
    End Property

#End Region

  End Class

End Namespace
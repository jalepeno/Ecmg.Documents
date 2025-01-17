'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

Namespace Core
  ''' <summary>Detailed information for the result of a search operation.</summary>
  <DebuggerDisplay("{ToString,nq")>
  Public Class SearchResult
    'Inherits Search.SearchResult
    Implements IComparable

#Region "Class Variables"

    Private mstrID As String
    Private mobjValues As New Data.DataItems

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpId As String)
      Me.New(lpId, New Data.DataItems)
    End Sub

    Public Sub New(ByVal lpId As String, ByVal lpValues As Data.DataItems)
      Try
        ID = lpId
        Values = lpValues
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Properties"

    Public Property ID() As String
      Get
        Return mstrID
      End Get
      Set(ByVal value As String)
        mstrID = value
      End Set
    End Property

    Public Property Values() As Data.DataItems
      Get
        Return mobjValues
      End Get
      Set(ByVal value As Data.DataItems)
        mobjValues = value
      End Set
    End Property

    Public ReadOnly Property IdValue As String
      Get
        Try
          Return GetIdValue()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property NameValue As String
      Get
        Try
          Return GetNameValue()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Public Methods"

    Public Overrides Function ToString() As String

      Try
        Dim lstrOutputString As String = ""

        ' Loop through each value and return the label and the value
        For Each lobjValue As Data.DataItem In Values
          lstrOutputString &= String.Format("{0}: {1}, ", lobjValue.Name, lobjValue.Value)
        Next

        lstrOutputString = lstrOutputString.TrimEnd(",", " ")

        Return lstrOutputString
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

#End Region

#Region "Private Methods"

    Private Function GetIdValue() As String
      Try
        ' Find the Id and return it
        If Values.Contains("Id") Then
          Return Values("Id").Value.ToString
        Else
          Return String.Empty
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function GetNameValue() As String
      Try
        ' Find the Id and return it
        If Values.Contains("Name") Then
          Return Values("Name").Value
        ElseIf Values.Contains("Title") Then
          Return Values("Title").Value
        ElseIf Values.Contains("DocumentTitle") Then
          Return Values("DocumentTitle").Value
        Else
          Return String.Empty
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IComparable Implementation"

    Public Function CompareTo(ByVal obj As Object) As Integer Implements IComparable.CompareTo

      Try
        If TypeOf obj Is SearchResult Then
          Return ID.CompareTo(obj.ID)
        Else
          Throw New ArgumentException("Object is not a SearchResult")
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

#End Region

  End Class

End Namespace
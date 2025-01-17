'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Text
Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Arguments

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class MigrateDocumentEventArgs
    Inherits DocumentEventArgs

#Region "Class Variables"

#End Region

#Region "Public Properties"

    Public Property OriginalId As String
    Public Property StartTime As DateTime
    Public Property CompleteTime As DateTime
    Public Property SynchronousMigration As Boolean
    Public Property Cancel As Boolean
    Public Property Record As Object

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpOriginalId As String)
      Me.New(lpOriginalId, New Document, Nothing, "Migrate Document", Now, DateTime.MinValue, True, False)
    End Sub

    Public Sub New(ByVal lpOriginalId As String,
                   ByVal lpStartTime As DateTime)
      Me.New(lpOriginalId, New Document, Nothing, "Migrate Document", lpStartTime, DateTime.MinValue, True, False)
    End Sub

    Public Sub New(ByVal lpOriginalId As String,
                   ByVal lpDocument As Document,
                   ByVal lpRecord As Object,
                   ByVal lpEventDescription As String,
                   ByVal lpStartTime As DateTime,
                   ByVal lpCompleteTime As DateTime,
                   ByVal lpSynchronousMigration As Boolean,
                   ByVal lpCancel As Boolean)
      MyBase.New(lpDocument, lpEventDescription, lpStartTime)
      OriginalId = lpOriginalId
      Record = lpRecord
      StartTime = lpStartTime
      CompleteTime = lpCompleteTime
      SynchronousMigration = lpSynchronousMigration
      Cancel = lpCancel
    End Sub

#End Region

#Region "Private Methods"

    Protected Friend Overrides Function DebuggerIdentifier() As String
      Try
        Dim lobjReturnBuilder As New StringBuilder

        If Cancel = True Then
          lobjReturnBuilder.Append("Cancel; ")
        End If

        If OriginalId.Length > 0 Then
          lobjReturnBuilder.AppendFormat("OriginalId={0}; ", OriginalId)
        End If

        lobjReturnBuilder.AppendFormat("{0}; ", MyBase.DebuggerIdentifier)

        lobjReturnBuilder.AppendFormat("Started {0}; ", StartTime)

        If CompleteTime > DateTime.MinValue Then
          lobjReturnBuilder.AppendFormat("Completed {0}; ", CompleteTime)
        End If

        If SynchronousMigration = True Then
          lobjReturnBuilder.Append("Synchronous; ")
        End If

        If Record IsNot Nothing Then
          lobjReturnBuilder.Append("Has record; ")
        End If

        lobjReturnBuilder.Remove(lobjReturnBuilder.Length - 2, 2)

        Return lobjReturnBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace

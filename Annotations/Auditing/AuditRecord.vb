'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

Namespace Annotations.Auditing

  ''' <summary>
  ''' Provides a record of strongly-typed events.
  ''' </summary>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class AuditRecord

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the creation user and time stamp information.
    ''' </summary>
    ''' <value>The user and time stamp information.</value>
    Public Property Created() As New CreateEvent

    ''' <summary>
    ''' Gets or sets the user and time stamp information of when the item was modified.
    ''' </summary>
    ''' <value>The user and time stamp information.</value>
    Public Property Modified() As New ModifyEvent

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Sets the creation informaion.
    ''' </summary>
    ''' <param name="created">The creation date.</param>
    ''' <param name="createUser">The user who created it.</param>
    Public Sub SetCreate(ByVal created As DateTime, ByVal createUser As String)
      Me.Created = New CreateEvent() With {.EventTime = created, .User = createUser}
    End Sub

    ''' <summary>
    ''' Sets the latest modification information.
    ''' </summary>
    ''' <param name="modified">The latest modification date.</param>
    ''' <param name="modifyUser">The user who modified it.</param>
    Public Sub SetModify(ByVal modified As DateTime, ByVal modifyUser As String)
      Me.Modified = New ModifyEvent() With {.EventTime = modified, .User = modifyUser}
    End Sub

#End Region

#Region "Private Methods"

    Protected Friend Function DebuggerIdentifier() As String
      Dim lobjIdentifierBuilder As New System.Text.StringBuilder
      Try

        lobjIdentifierBuilder.AppendFormat("Created: {0}, Modified: {1}", Created.DebuggerIdentifier, Modified.DebuggerIdentifier)

        Return lobjIdentifierBuilder.ToString

      Catch ex As System.Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
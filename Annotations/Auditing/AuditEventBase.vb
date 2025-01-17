'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports Documents.Utilities

#End Region

Namespace Annotations.Auditing

  <Serializable()>
  <XmlInclude(GetType(CreateEvent))>
  <XmlInclude(GetType(ModifyEvent))>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public MustInherit Class AuditEventBase

    Private utcEventTime As DateTime

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the user.
    ''' </summary>
    ''' <value>The user.</value>
    <XmlAttribute()>
    Public Property User() As String

    ''' <summary>
    ''' Gets or sets the event time.
    ''' </summary>
    ''' <value>The event time.</value>
    <XmlIgnore()>
    Public Property EventTime() As DateTime
      Get
        Return Me.utcEventTime.ToLocalTime()
      End Get
      Set(ByVal value As DateTime)
        Me.utcEventTime = value.ToUniversalTime()
      End Set
    End Property

    <XmlAttribute("EventTime")>
    Public Property EventTimeUTC() As DateTime
      Get
        Return Me.utcEventTime
      End Get
      Set(ByVal value As DateTime)
        Me.utcEventTime = value
      End Set
    End Property

#End Region

#Region "Private Methods"

    Protected Friend Function DebuggerIdentifier() As String
      Dim lobjIdentifierBuilder As New System.Text.StringBuilder
      Try

        If utcEventTime <> DateTime.MinValue Then
          If User IsNot Nothing AndAlso User.Length > 0 Then
            lobjIdentifierBuilder.AppendFormat("{0}: ", User)
          End If

          lobjIdentifierBuilder.AppendFormat("{0}", EventTime)
        Else
          lobjIdentifierBuilder.Append("Event not initialized")
        End If

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
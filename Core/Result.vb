'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports System.Xml.Serialization
Imports Documents.Utilities

Namespace Core

  ''' <summary>Base class containing detailed information for the result of an operation.</summary>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class ActionResult
    Implements IEquatable(Of ActionResult)

#Region "Class Variables"

    'Private mstrTestName As String
    Private mblnSuccess As Boolean
    Private mstrDetails As String = ""
    Private mobjStartTime As DateTime = Nothing
    Private mobjFinishTime As TimeSpan = Nothing

#End Region

#Region "Public Properties"

    'Public Property TestName() As String
    '  Get
    '    Return mstrTestName
    '  End Get
    '  Set(ByVal value As String)
    '    mstrTestName = value
    '  End Set
    'End Property

    Public Overridable Property Success() As Boolean
      Get
        Return mblnSuccess
      End Get
      Set(ByVal value As Boolean)
        mblnSuccess = value
      End Set
    End Property

    Public Property Details() As String
      Get
        Return mstrDetails
      End Get
      Set(ByVal value As String)
        mstrDetails = value
      End Set
    End Property

    <XmlAttribute("StartTime")>
    Public Property StartTime() As DateTime
      Get
        Return mobjStartTime
      End Get
      Set(ByVal value As DateTime)
        mobjStartTime = value
      End Set
    End Property

    <XmlIgnore()>
    Public Property ElapsedTime() As TimeSpan
      Get
        Return mobjFinishTime
      End Get
      Set(ByVal value As TimeSpan)
        mobjFinishTime = value
      End Set
    End Property

    <XmlAttribute("ElapsedTime")>
    Public Property ElapsedTimeString() As String
      Get
        Return ElapsedTime.ToString
      End Get
      Set(ByVal value As String)

      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpSuccess As Boolean)
      Success = lpSuccess
    End Sub

    Public Sub New(ByVal lpSuccess As Boolean, ByVal lpDetails As String)
      Success = lpSuccess
      Details = lpDetails
    End Sub

    Public Sub New(ByVal lpSuccess As Boolean, ByVal lpDetails As String,
                   ByVal lpStartTime As DateTime, ByVal lpElapsedTime As System.TimeSpan)

      Success = lpSuccess
      Details = lpDetails
      StartTime = lpStartTime
      ElapsedTime = lpElapsedTime

    End Sub

#End Region

#Region "Operator Overloads"

    Public Overloads Function Equals(ByVal lpResult As ActionResult) As Boolean Implements System.IEquatable(Of ActionResult).Equals
      Try
        If lpResult Is Nothing Then
          Return False
        ElseIf Me.Success = lpResult.Success Then
          Return True
        Else
          Return False
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Shared Operator =(ByVal lpResult As ActionResult, ByVal lpBoolean As Boolean) As Boolean

      Return lpResult.Success = lpBoolean

    End Operator

    Shared Operator <>(ByVal lpResult As ActionResult, ByVal lpBoolean As Boolean) As Boolean
      Return lpResult.Success <> lpBoolean
    End Operator

#End Region

#Region "Public Methods"

    Public Overridable Function ToResultCode() As Result
      Try
        If Success = True Then
          Return Result.Success
        Else
          Return Result.Failed
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    Friend Overridable Function DebuggerIdentifier() As String
      Try
        If Success = True Then
          If String.IsNullOrEmpty(Details) Then
            Return String.Format("Succeeded - {0}", ElapsedTimeString)
          Else
            Return String.Format("Succeeded ({0}) - {1}", Details, ElapsedTimeString)
          End If
        Else
          If String.IsNullOrEmpty(Details) Then
            Return String.Format("Failed {0}", ElapsedTimeString)
          Else
            Return String.Format("Failed ({0}) - {1}", Details, ElapsedTimeString)
          End If
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
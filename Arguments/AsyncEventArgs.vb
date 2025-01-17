'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------


Namespace Arguments

  ''' <summary>
  ''' Base class that will mimic AysncCompletedEventArgs(for web services)
  ''' </summary>
  ''' <remarks></remarks>
  Public Class AsyncEventArgs

    Private mobjResults() As Object
    Private mobjException As System.Exception
    Private mblnCancelled As Boolean = False
    Private mobjUserState As Object

    Public Sub New()
      ReDim mobjResults(0)
    End Sub

    Public Property Results() As Object
      Get
        Return mobjResults
      End Get
      Set(ByVal value As Object)
        mobjResults = value
      End Set
    End Property

    Public Property Exception() As Exception
      Get
        Return mobjException
      End Get
      Set(ByVal value As Exception)
        mobjException = value
      End Set
    End Property

    Public Property Cancelled() As Boolean
      Get
        Return mblnCancelled
      End Get
      Set(ByVal value As Boolean)
        mblnCancelled = value
      End Set
    End Property

    Public Property UserState() As Object
      Get
        Return mobjUserState
      End Get
      Set(ByVal value As Object)
        mobjUserState = value
      End Set
    End Property

  End Class

End Namespace

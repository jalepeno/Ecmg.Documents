'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Arguments

  ''' <summary>
  ''' Contains all the parameters necessary for the WriteMessage method
  ''' </summary>
  ''' <remarks></remarks>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class WriteMessageArgs
    Inherits EventArgs

#Region "Class Variables"
    Private mstrMessage As String
#End Region

#Region "Public Properties"

    Public ReadOnly Property Message() As String
      Get
        Return mstrMessage
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpMessage As String)
      mstrMessage = lpMessage
    End Sub

#End Region

#Region "Private Methods"

    Protected Friend Overridable Function DebuggerIdentifier() As String
      Try
        Return String.Format("Message={0}", Message)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
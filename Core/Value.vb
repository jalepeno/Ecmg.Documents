'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

Namespace Core
  ''' <summary>Contains the value of a single valued property.</summary>
  Partial Public Class Value

#Region "Class Variables"
    Private mobjValue As Object
#End Region

#Region "Public Properties"

    Public Property Value() As Object
      Get
        Return mobjValue
      End Get
      Set(ByVal Value As Object)
        mobjValue = Value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpValue As Object)
      Value = lpValue
    End Sub

    Public Sub New(ByVal lpValue As String)
      Value = lpValue
    End Sub

    Public Sub New(ByVal lpValue As Single)
      Value = lpValue
    End Sub

    Public Sub New(ByVal lpValue As Double)
      Value = lpValue
    End Sub

    Public Sub New(ByVal lpValue As Integer)
      Value = lpValue
    End Sub

    Public Sub New(ByVal lpValue As Long)
      Value = lpValue
    End Sub

    Public Sub New(ByVal lpValue As Boolean)
      Value = lpValue
    End Sub

    Public Sub New(ByVal lpValue As Date)
      Value = lpValue
    End Sub

#End Region

#Region "Operator Overloads"

    Public Overloads Function ToString() As String
      Try
        Return Me.Value.ToString()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

    Protected Overridable Function DebuggerIdentifier() As String
      Try
        Return Me.ToString()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

  End Class

End Namespace
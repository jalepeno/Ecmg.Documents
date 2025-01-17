'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  HeaderPattern.vb
'   Description :  [type_description_here]
'   Created     :  1/26/2015 1:59:04 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities
Imports Newtonsoft.Json

#End Region

Namespace Files

  <Serializable()> Public Class HeaderPattern
    Implements IHeaderPattern
    Implements IComparable

#Region "Class Variables"

    Private mstrBytesString As String = String.Empty
    Private mintPosition As Integer
    Private mintLength As Integer
    Private mobjPattern As Byte()
    Private mintPoints As Integer
    Private mblnXm As Boolean

#End Region

#Region "IHeaderPattern Implementation"

    <JsonProperty("bs")> Public Property Bytes As String Implements IHeaderPattern.Bytes
      Get
        Try
          Return mstrBytesString
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrBytesString = value
          InitializeFromByteString(value)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    <JsonProperty("pos")> Public Property Position As Integer Implements IHeaderPattern.Position
      Get
        Try
          Return mintPosition
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Integer)
        Try
          mintPosition = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Friend ReadOnly Property Length As Integer Implements IHeaderPattern.Length
      Get
        Try
          Return mintLength
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Friend ReadOnly Property Pattern As Byte() Implements IHeaderPattern.Pattern
      Get
        Try
          Return mobjPattern
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Friend ReadOnly Property Points As Integer Implements IHeaderPattern.Points
      Get
        Try
          Return mintPoints
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Friend ReadOnly Property XM As Boolean Implements IHeaderPattern.XM
      Get
        Try
          Return mblnXm
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Private Sub InitializeFromByteString(lpByteString As String)
      Try
        mintLength = (lpByteString.Length / 2)
        mobjPattern = New Byte((mintLength + 1) - 1) {}

        Dim lintCounter As Integer = 1
        Do While (lintCounter <= mintLength)
          mobjPattern((lintCounter - 1)) = CByte(Math.Round(Conversion.Val(("&H" & lpByteString.Substring(((lintCounter - 1) * 2), 2)))))
          lintCounter += 1
        Loop

        mblnXm = False

        If Position = 0 Then
          mintPoints = &H3E8
        Else
          mintPoints = 1
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "IComparable Implementation"

    Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
      Try
        Dim lintByteComparison As Integer = Bytes.CompareTo(obj.Bytes)
        If lintByteComparison <> 0 Then
          Return lintByteComparison
        End If

        Dim lintPositionComparison As Integer = Position.CompareTo(obj.Position)
        If lintPositionComparison <> 0 Then
          Return lintPositionComparison
        End If

        Return 0

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(lpBytes As String, lpPosition As Integer)
      Try
        Bytes = lpBytes
        Position = lpPosition
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace
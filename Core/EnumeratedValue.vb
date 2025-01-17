' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  EnumeratedValue.vb
'  Description :  [type_description_here]
'  Created     :  3/15/2012 9:46:05 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Text
Imports Documents.Utilities

#End Region

Namespace Core

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class EnumeratedValue
    Implements IEnumeratedValue

#Region "Class Variables"

    Private mstrParentName As String = String.Empty
    Private mstrName As String = String.Empty
    Private mintValue As Nullable(Of Integer)

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the name of the parent enumeration.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ParentName As String Implements IEnumeratedValue.ParentName
      Get
        Return mstrParentName
      End Get
      Set(value As String)
        mstrParentName = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the display name.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Name As String Implements IEnumeratedValue.Name
      Get
        Return mstrName
      End Get
      Set(value As String)
        mstrName = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the integer value.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Value As Nullable(Of Integer) Implements IEnumeratedValue.Value
      Get
        Return mintValue
      End Get
      Set(value As Nullable(Of Integer))
        mintValue = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(lpParentName As String, lpName As String, lpValue As Nullable(Of Integer))
      Try
        ParentName = lpParentName
        Name = lpName
        Value = lpValue
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Protected Methods"

    Protected Friend Overridable Function DebuggerIdentifier() As String
      Dim lobjIdentifierBuilder As New StringBuilder
      Try

        lobjIdentifierBuilder.AppendFormat("{0}: {1}({2})", ParentName, Name, Value)

        Return lobjIdentifierBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace
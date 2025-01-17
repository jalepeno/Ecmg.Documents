'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Exceptions

#End Region

Namespace Data.Exceptions

  Public Class ValueNotFoundException
    Inherits CtsException

#Region "Class Variables"

    Private mstrSQL As String

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' The SQL statement executed
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property SqlStatement() As String
      Get
        Return mstrSQL
      End Get
    End Property

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Creates a new ValueNotFoundException with the specified SQL statement.
    ''' </summary>
    ''' <param name="sqlStatement">The statement that was executed</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal sqlStatement As String)
      MyBase.New(String.Format("No value found for the expression ({0})", sqlStatement))
      mstrSQL = sqlStatement
    End Sub

    ''' <summary>
    ''' Creates a new ValueNotFoundException with the specified SQL statement.
    ''' </summary>
    ''' <param name="sqlStatement">The statement that was executed</param>
    ''' <param name="message">The exception message</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal sqlStatement As String,
                   ByVal message As String)
      MyBase.New(message)
      mstrSQL = sqlStatement
    End Sub

    ''' <summary>
    ''' Creates a new ValueNotFoundException with the specified SQL statement.
    ''' </summary>
    ''' <param name="sqlStatement">The statement that was executed</param>
    ''' <param name="message">The exception message</param>
    ''' <param name="innerException">Any inner exception to include</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal sqlStatement As String,
                   ByVal message As String,
                   ByVal innerException As System.Exception)
      MyBase.New(message, innerException)
      mstrSQL = sqlStatement
    End Sub

#End Region

  End Class

End Namespace

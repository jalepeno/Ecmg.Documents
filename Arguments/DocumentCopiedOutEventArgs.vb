'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Core

Namespace Arguments
  ''' <summary>
  ''' Contains all the parameters for the DocumentCopied Event
  ''' </summary>
  ''' <remarks></remarks>
  Public Class DocumentCopiedOutEventArgs
    Inherits DocumentEventArgs

#Region "Class Variables"

    Private mstrDestinationPath As String = ""
    Private mstrOutputFileNames As New List(Of String)

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' The directory the document was copied to.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property DestinationPath() As String
      Get
        Return mstrDestinationPath
      End Get
    End Property

    ''' <summary>
    ''' A string array containing the fully qualified path(s) to the files copied out 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Unless the repository supports multiple content elements per vesion 
    ''' and the version has multiple content elements, 
    ''' the array will only have one value populated
    ''' </remarks>
    Public ReadOnly Property OutputFileNames() As List(Of String)
      Get
        Return mstrOutputFileNames
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpDocument As Document, ByVal lpDestinationPath As String, ByVal lpOutputFileNames As String())
      Me.New(lpDocument, lpDestinationPath, lpOutputFileNames, Now)
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpDestinationPath As String, ByVal lpOutputFileNames As String(), ByVal lpTime As DateTime)
      MyBase.New(lpDocument, "DocumentCopiedOut", lpTime)
      mstrDestinationPath = lpDestinationPath
      For lintOutputFileNameCounter As Integer = 0 To lpOutputFileNames.Length - 1
        mstrOutputFileNames.Add(lpOutputFileNames(lintOutputFileNameCounter))
      Next
      'mstrOutputFileNames = lpOutputFileNames
    End Sub

#End Region

  End Class
End Namespace
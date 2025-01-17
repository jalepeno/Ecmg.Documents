' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ObjectPropertyArgs.vb
'  Description :  [type_description_here]
'  Created     :  2/29/2012 1:35:41 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Arguments

  Public Class ObjectPropertyArgs

#Region "Class Variables"

    Private mstrObjectID As String = String.Empty
    Private mobjProperties As IProperties = New ECMProperties

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Identifies the object to be updated
    ''' </summary>
    Public Property ObjectID() As String
      Get
        Return mstrObjectID
      End Get
      Set(ByVal value As String)
        mstrObjectID = value
      End Set
    End Property

    ''' <summary>
    ''' The properties to update in the document
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Each property needs to contain the values to be set in the document / version.</remarks>
    Public Property Properties() As IProperties
      Get
        Return mobjProperties
      End Get
      Set(ByVal value As IProperties)
        mobjProperties = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    ''' <summary>
    ''' Creates a new instance of ObjectPropertyArgs
    ''' </summary>
    ''' <param name="lpObjectId">The object identifier</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpObjectId As String)
      Me.New(lpObjectId, New ECMProperties)
    End Sub

    ''' <summary>
    ''' Creates a new instance of DocumentUpdateArgs
    ''' </summary>
    ''' <param name="lpObjectId">The document identifier</param>
    ''' <param name="lpProperties">A Core.Properties object containing the properties to update</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpObjectId As String, ByVal lpProperties As IProperties)

      Try

        ObjectID = lpObjectId
        Properties = lpProperties

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try

    End Sub

#End Region

  End Class

End Namespace
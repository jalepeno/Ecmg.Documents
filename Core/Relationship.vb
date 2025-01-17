'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

Namespace Core

  Public Class Relationship

#Region "Class Variables"

    Private mintOrder As Nullable(Of Integer)
    Private mstrID As String
    Private mobjRelationshipType As RelationshipType = RelationshipType.Child
    Private mobjRelationshipStrength As RelationshipStrength = RelationshipStrength.Strong
    Private menuRelationshipPersistance As RelationshipPersistance = RelationshipPersistance.Static
    Private mobjDocument As Document
    Private mstrParentId As String = String.Empty
    Private mstrChildId As String = String.Empty
    Private mstrLabelBindValue As String = String.Empty
    Private menuCascadeDeleteAction As RelationshipCascadeDeleteAction = RelationshipCascadeDeleteAction.NoCascadeDelete
    Private menuPreventDeleteAction As RelationshipPreventDeleteAction = RelationshipPreventDeleteAction.AllowBothDelete
    Private menuVersionBindType As VersionBindType = VersionBindType.LatestVersion
    Private mstrRelatedURI As String = String.Empty

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the order in which this relationship is ranked in the Relationships collection
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Order() As Nullable(Of Integer)
      Get
        Return mintOrder
      End Get
      Set(ByVal value As Nullable(Of Integer))
        mintOrder = value
      End Set
    End Property

    ''' <summary>
    ''' The unique identifier of the document or object to which the relationship belongs.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ObjectID() As String
      Get
        Return mstrID
      End Get
      Set(ByVal value As String)
        mstrID = value
      End Set
    End Property

    ''' <summary>
    ''' Defines hierarchy in a parent/child document relationship
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Type() As RelationshipType
      Get
        Return mobjRelationshipType
      End Get
      Set(ByVal value As RelationshipType)
        mobjRelationshipType = value
      End Set
    End Property

    ''' <summary>
    ''' Defines the strength of a document relationship
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Expressed as either strong or weak</remarks>
    Public Property Strength() As RelationshipStrength
      Get
        Return mobjRelationshipStrength
      End Get
      Set(ByVal value As RelationshipStrength)
        mobjRelationshipStrength = value
      End Set
    End Property

    ''' <summary>
    ''' Defines the spanning nature of a document relationship
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Expressed as either Dynamic or Static</remarks>
    Public Property Persistance() As RelationshipPersistance
      Get
        Return menuRelationshipPersistance
      End Get
      Set(ByVal value As RelationshipPersistance)
        ' mobjRelationshipPersistance = RelationshipPersistance.Dynamic
        menuRelationshipPersistance = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the flag that determines whether or not 
    ''' the related item should participate in a cascading delete.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CascadeDeleteAction As RelationshipCascadeDeleteAction
      Get
        Return menuCascadeDeleteAction
      End Get
      Set(ByVal value As RelationshipCascadeDeleteAction)
        menuCascadeDeleteAction = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the flag which determines whether or not a related item can be deleted.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PreventDeleteAction As RelationshipPreventDeleteAction
      Get
        Return menuPreventDeleteAction
      End Get
      Set(ByVal value As RelationshipPreventDeleteAction)
        menuPreventDeleteAction = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the value used for binding the relationship to the specific version of the related document.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LabelBindValue As String
      Get
        Return mstrLabelBindValue
      End Get
      Set(ByVal value As String)
        mstrLabelBindValue = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the value that determines whether a dynamic relationship 
    ''' is to bind with the latest version or the latest major version.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Only applies for repositories that can support this style of version binding.</remarks>
    Public Property VersionBindType As VersionBindType
      Get
        Return menuVersionBindType
      End Get
      Set(ByVal value As VersionBindType)
        menuVersionBindType = value
      End Set
    End Property

    ''' <summary>
    ''' The object specified for the relationship.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>If the Relationship Type is Child, then this property would refer to the child document.  If it is a Parent Relationship, then this property would refer to the parent document.</remarks>
    Public Property RelatedDocument() As Document
      Get
        Return mobjDocument
      End Get
      Set(ByVal value As Document)
        mobjDocument = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the URI for a URI relationship.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property RelatedURI As String
      Get
        Return mstrRelatedURI
      End Get
      Set(ByVal value As String)
        mstrRelatedURI = value
      End Set
    End Property

    Public Property ParentId As String
      Get
        Return mstrParentId
      End Get
      Set(ByVal value As String)
        mstrParentId = value
      End Set
    End Property

    Public ReadOnly Property ParentObjectId As String
      Get
        Return GetObjectId(ParentId)
      End Get
    End Property

    Public ReadOnly Property ParentVersionId As String
      Get
        Return GetVersionId(ParentId)
      End Get
    End Property

    Public Property ChildId As String
      Get
        Return mstrChildId
      End Get
      Set(ByVal value As String)
        mstrChildId = value
      End Set
    End Property

    Public ReadOnly Property ChildObjectId As String
      Get
        Return GetObjectId(ChildId)
      End Get
    End Property

    Public ReadOnly Property ChildVersionId As String
      Get
        Return GetVersionId(ChildId)
      End Get
    End Property

    Public Property ParentObjectType As RelatedObjectType

    Public Property ChildObjectType As RelatedObjectType

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal lpObjectID As String)
      MyBase.New()
      Try
        ObjectID = lpObjectID
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpObjectID As String,
                   ByVal lpParentID As String,
                   ByVal lpChildID As String,
                   ByVal lpParentType As RelatedObjectType,
                   ByVal lpChildType As RelatedObjectType,
                   ByVal lpType As RelationshipType,
                   ByVal lpStrength As RelationshipStrength,
                   ByVal lpPersistance As RelationshipPersistance)
      MyBase.New()
      Try
        ObjectID = lpObjectID
        ParentId = lpParentID
        ParentObjectType = lpParentType
        ChildObjectType = lpChildType
        ChildId = lpChildID
        Type = lpType
        Strength = lpStrength
        Persistance = lpPersistance
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Overridable Function IsValid(ByVal lpParentDocument As Document) As Boolean
      Try

        If ParentObjectType = RelatedObjectType.Document AndAlso
          ChildObjectType = RelatedObjectType.Document Then
          Return True
        End If

        Select Case Type
          Case RelationshipType.Parent
            If lpParentDocument.Versions.IdExists(ParentVersionId) AndAlso
              RelatedDocument.Versions.IdExists(ChildVersionId) Then
              ' Both the versions are still there, the relationship is valid.
              Return True
            Else
              ' They are not both there, the relationship is not valid.
              Return False
            End If
          Case RelationshipType.Child
            If lpParentDocument.Versions.IdExists(ChildVersionId) AndAlso
                RelatedDocument.Versions.IdExists(ParentVersionId) Then
              ' Both the versions are still there, the relationship is valid.
              Return True
            Else
              ' They are not both there, the relationship is not valid.
              Return False
            End If
          Case Else
            Return False
        End Select

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    Protected Function GetObjectId(ByVal lpRelatedDocId As String) As String
      Try
        Dim lstrObjectId As String = Nothing
        SplitRelatedDocId(lpRelatedDocId, lstrObjectId, Nothing)
        Return lstrObjectId
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Function GetVersionId(ByVal lpRelatedDocId As String) As String
      Try
        Dim lstrVersionId As String = Nothing
        SplitRelatedDocId(lpRelatedDocId, Nothing, lstrVersionId)
        Return lstrVersionId
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Splits the related document id into separate parts for the object id and the version id.
    ''' </summary>
    ''' <param name="lpRelatedDocId">The related document id to split.</param>
    ''' <param name="lpObjectId">The resulting object id.</param>
    ''' <param name="lpVersionId">The resulting version id.</param>
    ''' <remarks>Assumes that they are delimited by a dash '-'</remarks>
    Protected Sub SplitRelatedDocId(ByVal lpRelatedDocId As String,
                                    ByRef lpObjectId As String,
                                    ByRef lpVersionId As String)
      Try

        If String.IsNullOrEmpty(lpRelatedDocId) Then
          Throw New ArgumentNullException("lpRepatedDocumentId")
        End If

        Dim lstrIdParts() As String = lpRelatedDocId.Split("-")

        If lstrIdParts.Length = 2 Then
          lpObjectId = lstrIdParts(0)
          lpVersionId = lstrIdParts(1)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  FolderClass.vb
'  Description :  [type_description_here]
'  Created     :  3/5/2012 8:53:23 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"


#End Region

Namespace Core

  Public Class FolderClass
    Inherits RepositoryObjectClass

#Region "Constructors"

    ''' <summary>Default Constructor</summary>
    Public Sub New()
      MyBase.New()
    End Sub

    ''' <summary>
    ''' Constructs a new folder class object using the specified name, properties
    ''' collection, ID and label.
    ''' </summary>
    Public Sub New(ByVal lpName As String,
                   ByVal lpProperties As ClassificationProperties,
                   Optional ByVal lpId As String = "",
                   Optional ByVal lpLabel As String = "")
      MyBase.New(lpName, lpProperties, lpId, lpLabel)
    End Sub

    ''' <summary>
    ''' Constructs a new folder class object using the specified stream.
    ''' </summary>
    ''' <param name="lpStream">An IO.Stream object derived from a folder class file.</param>
    ''' <remarks>
    ''' If there are child choice lists, they can't be retrieved 
    ''' using this constructor.  Use New(lpStream, lpZipFile) instead.
    '''  The choice lists are actually stored in separate files that 
    ''' can only be from the zipped repository archive file.
    '''</remarks>
    Public Sub New(ByVal lpStream As IO.Stream)
      MyBase.New(lpStream)
    End Sub

#End Region

  End Class

End Namespace
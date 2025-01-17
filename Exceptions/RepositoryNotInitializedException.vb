' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  RepositoryNotInitializedException.vb
'  Description :  To be thrown when a repository object is needed but not yet 
'                 initialized.
'  Created     :  11/9/2011 11:01:56 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"



#End Region

Namespace Exceptions

  ''' <summary>
  ''' To be thrown when a repository object is needed but not yet initialized.
  ''' </summary>
  ''' <remarks></remarks>
  Public Class RepositoryNotInitializedException
    Inherits CtsException

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal message As String)
      MyBase.New(message)
    End Sub

    Public Sub New(ByVal message As String,
                   ByVal innerException As Exception)
      MyBase.New(message, innerException)
    End Sub

#End Region

  End Class

End Namespace
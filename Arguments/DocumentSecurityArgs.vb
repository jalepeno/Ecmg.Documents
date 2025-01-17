' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  DocumentSecurityArgs.vb
'  Description :  [type_description_here]
'  Created     :  3/21/2012 3:58:44 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Security

#End Region

Namespace Arguments

  Public Class DocumentSecurityArgs
    Inherits ObjectSecurityArgs

#Region "Constructors"

    ''' <summary>
    ''' Creates a new instance of DocumentSecurityArgs
    ''' </summary>
    ''' <remarks>This is the default constructor primarily used for serialization.</remarks>
    Public Sub New()
      MyBase.New()
    End Sub

    ''' <summary>
    ''' Creates a new instance of DocumentSecurityArgs
    ''' </summary>
    ''' <param name="lpObjectId">The document identifier</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpObjectId As String)
      MyBase.New(lpObjectId)
    End Sub

    ''' <summary>
    ''' Creates a new instance of DocumentSecurityArgs
    ''' </summary>
    ''' <param name="lpObjectId">The document identifier</param>
    ''' <param name="lpPermissions">A IPermissions object containing the permissions to update</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpObjectId As String, ByVal lpPermissions As IPermissions)
      MyBase.New(lpObjectId, lpPermissions)
    End Sub

#End Region

  End Class

End Namespace
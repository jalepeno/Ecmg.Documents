'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  ImportObjectArgs.vb
'   Description :  [type_description_here]
'   Created     :  9/2/2015 12:30:11 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Arguments

  Public Class ImportObjectArgs
    Inherits BackgroundWorkerEventArgs

#Region "Class Variables"

    Private mobjObject As CustomObject
    Private mblnSetPermissions As Boolean = True

#End Region

#Region "Public Properties"

    Public Property [Object] As CustomObject
      Get
        Try
          Return mobjObject
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As CustomObject)
        Try
          mobjObject = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>Determines whether or not the permissions will be set on the destination folder using the permissions collection defined in the source folder.</summary>
    ''' <remarks>This property can only be used on imports with providers that implement IUpdatePermissions.</remarks>
    Public Property SetPermissions As Boolean
      Get
        Return mblnSetPermissions
      End Get
      Set(value As Boolean)
        mblnSetPermissions = value
      End Set
    End Property

#End Region

  End Class

End Namespace
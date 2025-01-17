'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  ContentSourceEventArgs.vb
'   Description :  [type_description_here]
'   Created     :  1/4/2013 3:46:06 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Providers
Imports Documents.Utilities


#End Region

Namespace Arguments

  Public Class ContentSourceEventArgs
    Inherits ItemEventArgs

#Region "Class Variables"

#End Region

#Region "Public Properties"

    Public Overloads ReadOnly Property Item As ContentSource
      Get
        Try
          Return MyBase.Item
        Catch Ex As Exception
          ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(contentSource As ContentSource)
      MyBase.New(contentSource)
    End Sub

#End Region

  End Class

End Namespace
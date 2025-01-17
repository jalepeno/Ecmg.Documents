' ********************************************************************************
' '  Document    :  ObjectIdentifiers.vb
' '  Description :  [type_description_here]
' '  Created     :  11/13/2012-16:04:15
' '  <copyright company="ECMG">
' '      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
' '      Copying or reuse without permission is strictly forbidden.
' '  </copyright>
' ********************************************************************************

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Core

  Public Class ObjectIdentifiers
    Inherits CCollection(Of ObjectIdentifier)

    Public Overloads Sub Add(ByVal lpIdentifierValue As String,
                           ByVal lpObjectType As ObjectIdentifier.ObjectTypeEnum)
      Try
        MyBase.Add(New ObjectIdentifier(lpIdentifierValue, lpObjectType))
      Catch Ex As Exception
        ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Sub Add(ByVal lpIdentifierValue As String,
                             ByVal lpIdentifierType As ObjectIdentifier.IdTypeEnum,
                             ByVal lpObjectType As ObjectIdentifier.ObjectTypeEnum)
      Try
        MyBase.Add(New ObjectIdentifier(lpIdentifierValue, lpIdentifierType, lpObjectType))
      Catch Ex As Exception
        ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

  End Class

End Namespace
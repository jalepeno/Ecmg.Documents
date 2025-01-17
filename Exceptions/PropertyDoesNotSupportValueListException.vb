'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core

#End Region

Namespace Exceptions

  Public Class PropertyDoesNotSupportValueListException
    Inherits PropertyException

#Region "Constructors"

    ''' <summary>
    ''' Creates a new PropertyDoesNotSupportValueListException
    ''' </summary>
    ''' <param name="property">The property that is causing the exception</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal [property] As IProperty)
      MyBase.New(String.Format("Property '{0}' does not support a value list.", [property].Name))
    End Sub

    ''' <summary>
    ''' Creates a new PropertyDoesNotSupportValueListException
    ''' </summary>
    ''' <param name="message">The message to include for the exception</param>
    ''' <param name="property">The property that is causing the exception</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal message As String, ByVal [property] As IProperty)
      MyBase.New(message, [property])
    End Sub

#End Region

  End Class

End Namespace

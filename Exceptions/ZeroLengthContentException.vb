' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ZeroLengthContentException.vb
'  Description :  Used in cases where the content is empty.
'  Created     :  9/27/2011 2:32:17 PM
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
  ''' To be thrown when we are requested to create content with a zero length.
  ''' </summary>
  ''' <remarks></remarks>
  Public Class ZeroLengthContentException
    Inherits CtsException

#Region "Constructors"

    Public Sub New()
      MyBase.New("Content has zero length.")
    End Sub

    Public Sub New(ByVal message As String)
      MyBase.New(message)
    End Sub

    Public Sub New(ByVal message As String, ByVal innerException As Exception)
      MyBase.New(message, innerException)
    End Sub

#End Region

  End Class

End Namespace
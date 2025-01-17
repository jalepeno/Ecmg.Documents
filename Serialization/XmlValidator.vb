'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  XmlValidator.vb
'   Description :  [type_description_here]
'   Created     :  1/15/2013 8:56:31 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization

#End Region

Namespace Serialization

  ''' <summary>
  '''     Used to validate xml object representations prior to deserializing
  ''' </summary>
  ''' <remarks>
  '''     
  ''' </remarks>
  Public Class XmlValidator
    Inherits Validator

    ''' <summary>
    '''     Validates an XML string for deserializing a specific object type.
    ''' </summary>
    ''' <param name="lpXml" type="String">
    '''     <para>
    '''         
    '''     </para>
    ''' </param>
    ''' <param name="lpType" type="System.Type">
    '''     <para>
    '''         
    '''     </para>
    ''' </param>
    ''' <param name="lpErrorMessage" type="String">
    '''     <para>
    '''         
    '''     </para>
    ''' </param>
    ''' <returns>
    '''     A Boolean value...
    ''' </returns>
    Public Shared Function ValidateXmlString(ByVal lpXml As String, ByVal lpType As Type, ByRef lpErrorMessage As String) As Boolean
      Dim lobjReader As System.IO.StringReader = Nothing
      Try
        Dim lobjFormatter As New XmlSerializer(lpType)
        lobjReader = New System.IO.StringReader(lpXml)
        Dim lobjTargetObject As Object = lobjFormatter.Deserialize(lobjReader)
        lobjTargetObject = Nothing
        Return True
      Catch ex As Exception
        lpErrorMessage = ex.Message
        Return False
      Finally
        If lobjReader IsNot Nothing Then lobjReader.Close()
      End Try
    End Function

  End Class

End Namespace
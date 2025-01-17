'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  IActionItem.vb
'   Description :  [type_description_here]
'   Created     :  4/24/2014 10:54:34 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"



#End Region

Namespace Core

  Public Interface IActionItem

    ReadOnly Property Name As String

    ReadOnly Property DisplayName As String

    ReadOnly Property Description As String

    ''' <summary>
    ''' Used to specify instructions or parameters.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property Parameters As IParameters

    ''' <summary>
    '''     Gets the value of the specified parameter if it exists, otherwise returns the default value.
    ''' </summary>
    ''' <param name="lpParameterName" type="String">
    '''     <para>
    '''         The name of the parameter to resolve.
    '''     </para>
    ''' </param>
    ''' <param name="lpDefaultValue" type="Object">
    '''     <para>
    '''         The default value to return if the parameter is not available.
    '''     </para>
    ''' </param>
    ''' <returns>
    '''     Returns the specified parameter if it exists, otherwise returns the default value.
    ''' </returns>
    Function GetParameterValue(ByVal lpParameterName As String, ByVal lpDefaultValue As Object) As Object
    Function GetEnumParameterValue(ByVal lpParameterName As String, ByVal lpEnumType As Type, ByVal lpDefaultValue As Object) As [Enum]
    Function GetStringParameterValue(ByVal lpParameterName As String, ByVal lpDefaultValue As Object) As String
    Function GetBooleanParameterValue(ByVal lpParameterName As String, ByVal lpDefaultValue As Object) As Boolean

    Sub SetDescription(ByVal lpDescription As String)

    Function ToXmlElementString() As String

  End Interface

End Namespace
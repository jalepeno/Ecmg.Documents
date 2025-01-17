' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IActionableBase.vb
'  Description :  [type_description_here]
'  Created     :  11/2/2013 12:56:41 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Globalization

#End Region

Namespace Core

  Public Interface IActionableBase
    Inherits IDisposable
    Inherits ICloneable


    ReadOnly Property Name As String

    ReadOnly Property DisplayName As String

    ReadOnly Property Description As String

    ''' <summary>
    ''' The Id of the document to operate on.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property DocumentId As String

    ''' <summary>
    ''' Used to specify instructions or parameters.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property Parameters As IParameters

    ' ''' <summary>
    ' ''' The batch with which to execute the item.
    ' ''' </summary>
    ' ''' <value></value>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Property Batch As Batch

    Property ActionableParent As IActionableBase


    ''' <summary>
    ''' Used to pass any message regarding the execution of the operation.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property ProcessedMessage As String

    Property ShouldExecute As Boolean

    ''' <summary>
    ''' Used to indicate the result.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ReadOnly Property Result As Result

    ''' <summary>
    ''' Optional property used to provide a reference to the host 
    ''' application or form under which the operation or proces is running.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>Property Host As Object
    Property Host As Object

    ''' <summary>
    ''' Optional tag which can be associated with the operation or process.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property Tag As Object

    ReadOnly Property IsDisposed As Boolean

    ReadOnly Property Locale As CultureInfo

    Sub SetDescription(ByVal lpDescription As String)
    Sub SetResult(ByVal lpResult As Result)

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

    Function OnExecute() As Result

    ReadOnly Property InstanceId As String

    Sub SetInstanceId(lpInstanceId As String)

    Function ToString() As String
    Function ToXmlElementString() As String

  End Interface

End Namespace
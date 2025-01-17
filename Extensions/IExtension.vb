' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IExtension.vb
'  Description :  Used for managing extension dlls.
'  Created     :  11/16/2011 7:35:43 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"


#End Region

Namespace Extensions

  Public Interface IExtension
    Inherits IExtensionInformation
    Inherits IDisposable

#Region "Properties"

    ''' <summary>
    ''' Gets or sets the identifier for the extension.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property Id As String

    ''' <summary>
    ''' Gets or sets a byte array containing the extension assembly.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property ByteArray As Byte()

    ''' <summary>
    ''' Gets or sets the date the extension was added.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property CreateDate As Date

    ''' <summary>
    ''' Gets or sets the user that added the extension.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property AddedByUser As String

    ''' <summary>
    ''' Gets or sets the machine name that the extension was originally added by.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property AddedByMachine As String

#End Region

  End Interface

End Namespace
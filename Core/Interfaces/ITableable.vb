' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ITableable.vb
'  Description :  Supports representing the object as a DataTable.
'  Created     :  6/20/2012 3:07:41 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

Imports System.Data

Namespace Core

  ''' <summary>
  ''' Supports representing the object as a DataTable.
  ''' </summary>
  ''' <remarks></remarks>
  Public Interface ITableable

    ''' <summary>
    ''' Returns a representation of the current object as a DataTable
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function ToDataTable() As DataTable

  End Interface

End Namespace
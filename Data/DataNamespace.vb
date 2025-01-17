'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"




#End Region

''' <summary>
''' Contains objects used for interacting with databases
''' </summary>
Namespace Data

  ''' <summary>Describes the type of QueryTarget that will be searched.</summary>
  Public Enum QueryTargetType
    Table = 0
    View = 1
    SQLStatement = 2
  End Enum
  Public Enum SortDirection
    Desc = 0
    Asc = 1
  End Enum
  Public Enum SearchType
    PropertySearch = 0
    ContentSearch = 1
  End Enum

End Namespace

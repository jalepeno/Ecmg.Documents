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
Namespace Search

  '''' <summary>Describes the type of QueryTarget that will be searched.</summary>
  'Public Enum QueryTargetType
  '  Table = 0
  '  View = 1
  '  SQLStatement = 2
  'End Enum
  Public Enum SortDirection
    Desc = 0
    Asc = 1
  End Enum

  Public Enum SearchType
    PropertySearch = 0
    ContentSearch = 1
  End Enum

  ''' <summary>
  ''' Defines how multiple criteria should be evaluated relative to each other for building searches.
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum SetEvaluation
    seAnd = 0
    seOr = 1
  End Enum

  ''' <summary>
  ''' Defines how a criterion should be displayed for search screens
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum DisplayType
    ''' <summary>
    ''' Do not show
    ''' </summary>
    ''' <remarks></remarks>
    Hidden = 0
    ''' <summary>
    ''' Show and require a value
    ''' </summary>
    ''' <remarks></remarks>
    Required = 1
    ''' <summary>
    ''' Show and allow the user to choose 
    ''' whether or not to provide a value
    ''' </summary>
    ''' <remarks></remarks>
    Editable = 2
  End Enum

  ''' <summary>
  ''' Defines how a field is to be sorted.
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum SortOption
    ''' <summary>
    ''' Not Sorted
    ''' </summary>
    ''' <remarks></remarks>
    None = 0
    ''' <summary>
    ''' Sort in ascending order
    ''' </summary>
    ''' <remarks></remarks>
    Ascending = 1
    ''' <summary>
    ''' Sort in descending order
    ''' </summary>
    ''' <remarks></remarks>
    Descending = 2
  End Enum

  ''' <summary>
  ''' Defines how an item is to be aligned
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum AlignmentOption
    ''' <summary>
    ''' Left aligned
    ''' </summary>
    ''' <remarks></remarks>
    Left = 0
    ''' <summary>
    ''' Center aligned
    ''' </summary>
    ''' <remarks></remarks>
    Center = 1
    ''' <summary>
    ''' Right aligned
    ''' </summary>
    ''' <remarks></remarks>
    Right = 2
    ''' <summary>
    ''' Stretched or justified
    ''' </summary>
    ''' <remarks></remarks>
    Stretch = 3
  End Enum

#Region "Enumerations"

  Public Enum FilterView
    [ReadOnly] = 1
    Editable = 0
    Required = 2
    Hidden = 3
  End Enum

#End Region

End Namespace

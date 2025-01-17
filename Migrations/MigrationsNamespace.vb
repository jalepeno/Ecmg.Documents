'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------


''' <summary>
''' Contains objects used in migration operations
''' </summary>
Namespace Migrations

#Region "Public Enums"

  ''' <summary>
  ''' Describes whether the path component will be placed in the front of or the back of
  ''' the existing path.
  ''' </summary>
  Public Enum ePathLocation
    Front = 0
    Back = 1
  End Enum

  ''' <summary>Describes the scope of a package.</summary>
  Public Enum PackageType
    Documents = 1
    Folder = 2
    SearchResults = 3
    EntireContentSource = 4
  End Enum


  Public Enum ProcessedStatus
    NotProcessed = 0
    Success = 1
    Failed = 2
    Processing = 3
  End Enum

  Public Enum OperationType
    Export = 0
    Migrate = 1
    Delete = 2
    CheckIn = 3
    CheckOut = 4
    CancelCheckOut = 5
  End Enum

#End Region

End Namespace
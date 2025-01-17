' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  PermissionType.vb
'  Description :  [type_description_here]
'  Created     :  3/15/2012 11:10:38 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"




#End Region

Namespace Security

  Public Enum AccessType
    Allow = 1
    Deny = 2
  End Enum

  Public Enum PermissionSource
    Direct = 0
    [Default] = 1
    Template = 2
    Parent = 3
    Marking = 4
    Proxy = 255
  End Enum

  Public Enum PermissionType
    Level = 0
    DefaultLevel = 1
    Right = 2
    InheritOnlyRight = 3
  End Enum

  Public Enum PrincipalType
    Unknown = 0
    User = 2000
    Group = 2001
  End Enum

  Public Enum PermissionLevel
    Custom
    FullControl
    ModifyProperties
    AddToFolder
    ViewProperties
    AllButOwnerControlDocumentRollup
  End Enum

  Public Enum PermissionRight
    ViewDocumentProperties
    ViewFolderProperties
    ModifyDocumentProperties
    ModifyDocumentPropertiesRollup
    ModifyFolderPropertiesRollup
    ViewContent
    ViewContentRollup
    ModifyContentRollup
    FileInFolder
    Link_Annotate
    PromoteVersionRollup
    Publish
    PublishRollup
    CreateInstance
    ChangeState
    MinorVersioning
    MajorVersioning
    Delete
    ReadPermissions
    ModifyPermissions
    ModifyOwner
    UnlinkDocument
    UnfileFromFolder
    CreateSubfolder
    CreateSubfolderRollup
    FullControlRollup
  End Enum

End Namespace

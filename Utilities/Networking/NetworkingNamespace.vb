'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  NetworkingNamespace.vb
'   Description :  [type_description_here]
'   Created     :  1/3/2013 3:54:51 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

Namespace Network

#Region "Public Enumerations"

  Public Enum ProtocolEnum
    file = 0
    ftp = 1
    gopher = 2
    http = 3
    https = 4
    ldap = 5
    mailto = 6
    pipe = 7
    tcp = 8
    news = 9
    nntp = 10
    sftp = 11
    telnet = 12
    uuid = 13
  End Enum

#End Region

End Namespace
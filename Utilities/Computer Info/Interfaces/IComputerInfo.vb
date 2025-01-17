'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  IComputerInfo.vb
'   Description :  [type_description_here]
'   Created     :  1/9/2014 8:01:30 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

Namespace Computer

  Public Interface IComputerInfo

    ReadOnly Property OperatingSystem As OperatingSystemInfo

    ReadOnly Property Processor As ProcessorInfo

    Function ToJson() As String

    Function ToStatusJson() As String

  End Interface

End Namespace
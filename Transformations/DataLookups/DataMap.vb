'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Data

Namespace Transformations
  ''' <summary>
  ''' Executes a database search using the specified criteria and applies the the
  ''' return value to the destination property as the new value.
  ''' </summary>
  ''' <remarks>Pass though inherited class of Cts.DatabaseUtilities.DataSource.</remarks>
  <Serializable()>
  Public Class DataMap
    Inherits DataSource

  End Class
End Namespace
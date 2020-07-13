param(
  $SiteUrl = 'https://contoso.sharepoint.com/sites/Footer',
  $SiteSponsorSecurityGroupName = 'SiteSponsorEditors'
)

Connect-PnPOnline -Url $SiteUrl

### Site Configuration List ###
$ConfigListTitle = "SiteConfig"

$list = Get-PnPList -Identity $ConfigListTitle

if (-not $list) {
  # Create List if it doesn't exist
  $list = New-PnPList -Title $ConfigListTitle -Template "GenericList" -Hidden -EnableVersioning

  # Hide List from Search
  $list.NoCrawl = $true
  $list.Update()
  $list.Context.ExecuteQuery()

  # Create/Update List Columns
  Add-PnPField -List $ConfigListTitle -DisplayName 'Value' -InternalName 'Value' -Type Note -AddToDefaultView | Out-Null
  Set-PnPField -List $ConfigListTitle -Identity Title -Values @{"EnforceUniqueValues" = $True; "Indexed" = $True}

  # Break List Permissions
  Set-PnPList -Identity $ConfigListTitle -BreakRoleInheritance -CopyRoleAssignments

  # Reduce Members group to Read-only
  Set-PnPListPermission -Identity $ConfigListTitle -Group (Get-PnPGroup -AssociatedMemberGroup) -AddRole "Read" -RemoveRole "Edit"

  # Create Site Sponsor List Item
  $listItem = Add-PnPListItem -List $ConfigListTitle -Values @{"Title" = "SITE_SPONSOR"; "Value" = ""}

  # Set Site Sponsor Item Permissions to allow Site Sponsor Security Group to edit
  Set-PnPListItemPermission -List $ConfigListTitle -Identity $listItem -User $SiteSponsorSecurityGroupName -AddRole "Contribute"

  # Create Site Primary Administrator List Item
  Add-PnPListItem -List $ConfigListTitle -Values @{"Title" = "SITE_PRIMARY_ADMIN"; "Value" = ""} | Out-Null

}



# To target all non-Site Owner assignments on list
# $list = Get-PnPList -Identity $ConfigListTitle
# $readRole = Get-PnPRoleDefinition -Identity Read
# $roleAssignments = Get-PnPProperty -ClientObject $list -Property RoleAssignments

# foreach ($ra in $roleAssignments) {
#   $bindings = Get-PnPProperty -ClientObject $ra -Property RoleDefinitionBindings
#   foreach ($bi in $bindings) {
#     if ($bi.RoleTypeKind -ne "Administrator" -and $bi.RoleTypeKind -ne "Reader") {
#       $ra.RoleDefinitionBindings.Remove($bi)
#       $ra.RoleDefinitionBindings.Add($readRole)
#     }
#   }
#   $ra.Update()
#   $ra.Context.ExecuteQuery()
# }

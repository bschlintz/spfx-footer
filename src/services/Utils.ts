export const toUpn = (loginName: string): string => loginName ? loginName.replace('i:0#.f|membership|', '') : '';

export const getWebTemplateFriendlyName = (webTemplate: string): string => {
  switch (webTemplate)
  {
    case "STS#3":                     return "Team site (no Office 365 group)";
    case "STS#0":                     return "Team site (classic experience)";
    case "BDR#0":                     return "Document Center";
    case "DEV#0":                     return "Developer Site";
    case "OFFILE#1":                  return "Team Site - SharePoint Online configuration";
    case "BICenterSite#0":            return "Business Intelligence Center";
    case "SRCHCEN#0":                 return "Enterprise Search Center";
    case "BLANKINTERNETCONTAINER#0":  return "Publishing Portal";
    case "ENTERWIKI#0":               return "Enterprise Wiki";
    case "PROJECTSITE#0":             return "Project Site";
    case "PRODUCTCATALOG#0":          return "Product Catalog";
    case "COMMUNITY#0":               return "Community Site";
    case "COMMUNITYPORTAL#0":         return "Community Portal";
    case "SITEPAGEPUBLISHING#0":      return "Communication site";
    case "SRCHCENTERLITE#0":          return "Basic Search Center";
    case "visprus#0":                 return "Visio Process Repository";
    case "GROUP#0":                   return "Team site (Office 365 group)";
    default:                          return "Other";
  }
};

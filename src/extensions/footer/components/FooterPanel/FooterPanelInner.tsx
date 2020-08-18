import * as React from 'react';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import * as strings from 'FooterApplicationCustomizerStrings';
import { SiteService } from '../../../../services/SiteService';
import SponsorField from '../SponsorField/SponsorField';
import DisplayField from '../DisplayField/DisplayField';
import SPPermission from '@microsoft/sp-page-context/lib/SPPermission';
import useAsyncData from '../../../../hooks/useAsyncData';
import { SiteStats } from '../../../../models/SiteStats';
import Person from '../Person/Person';
import { SiteUser } from '../../../../models/SiteUser';
import { WebStats } from '../../../../models/WebStats';
import { getWebTemplateFriendlyName } from '../../../../services/Utils';

export interface IFooterPanelInnerProps {
  siteService: SiteService;
  supportLink: string;
}

const FooterPanelInner: React.FC<IFooterPanelInnerProps> = ({ siteService, supportLink }) => {
  const isCurrentUserAnAdmin = siteService.spfxContext.pageContext.web.permissions.hasPermission(SPPermission.manageWeb as any);
  const isOffice365Group = !!siteService.spfxContext.pageContext.site.group;

  const {
    data: primaryAdmin,
    isLoading: isLoadingPrimaryAdmin
  } = useAsyncData<SiteUser>(null, siteService.getPrimaryAdmin, []);

  const {
    data: siteAdminsOrGroupOwners,
    isLoading: isLoadingSiteAdminsOrGroupOwners
  } = useAsyncData<SiteUser[]>(null, isCurrentUserAnAdmin ? siteService.getSiteAdminsOrGroupOwners : () => {}, []);

  const {
    data: siteStats,
    isLoading: isLoadingSiteStats
  } = useAsyncData<SiteStats>(null, isCurrentUserAnAdmin ? siteService.getSiteStats : () => {}, []);

  const {
    data: webStats,
    isLoading: isLoadingWebStats
  } = useAsyncData<WebStats>(null, siteService.getWebStats, []);

  return (
    <Stack tokens={{ childrenGap: 12 }} style={{ flex: 1 }}>
      {/* Sponsor */}
      <SponsorField siteService={siteService} />

      {/* Primary Site Admin */}
      <DisplayField label={strings.PrimarySiteAdminLabel} isLoading={isLoadingPrimaryAdmin}>
        {primaryAdmin
          ? <Person user={primaryAdmin} siteService={siteService} />
          : <span>{strings.EmptyFieldLabel}</span>
        }
      </DisplayField>

      {/* Site Administrators or Group Owners */}
      <DisplayField
        label={isOffice365Group ? strings.GroupOwnersLabel : strings.SiteAdminsLabel}
        hidden={isLoadingSiteAdminsOrGroupOwners || !siteAdminsOrGroupOwners || (siteAdminsOrGroupOwners && siteAdminsOrGroupOwners.length === 0)}
      >
        <Stack tokens={{ childrenGap: 6 }}>
          {siteAdminsOrGroupOwners && siteAdminsOrGroupOwners.map(sa => <Person user={sa} siteService={siteService} />)}
        </Stack>
      </DisplayField>

      {/* Storage */}
      <DisplayField label={strings.StorageLabel} isLoading={isLoadingSiteStats} hidden={!isCurrentUserAnAdmin}>
        {siteStats && siteStats.storageUsedBytes > 0
          ? <span>{(siteStats.storageUsedBytes / 1024 / 1024).toFixed(2)} MB used, {Math.floor(100 - siteStats.storageUsedPercentage)}% free</span>
          : <span>{strings.NoDataFieldLabel}</span>
        }
      </DisplayField>

      {/* Created */}
      <DisplayField label={strings.CreatedLabel} isLoading={isLoadingWebStats}>
        {webStats && webStats.created
          ? <span>{`${webStats.created.toLocaleString()}`}</span>
          : <span>{strings.NoDataFieldLabel}</span>
        }
      </DisplayField>

      {/* Last Updated */}
      <DisplayField label={strings.LastUpdatedLabel} isLoading={isLoadingWebStats}>
        {webStats && webStats.lastUpdated
          ? <span>{`${webStats.lastUpdated.toLocaleString()}`}</span>
          : <span>{strings.NoDataFieldLabel}</span>
        }
      </DisplayField>

      {/* Web Template */}
      <DisplayField label={strings.SiteTemplateLabel} isLoading={isLoadingWebStats}>
        {webStats && webStats.webTemplate
          ? <span>{getWebTemplateFriendlyName(webStats.webTemplate)}</span>
          : <span>{strings.NoDataFieldLabel}</span>
        }
      </DisplayField>

      {/* Office 365 Group ID */}
      <DisplayField label={strings.GroupIdLabel} hidden={!isOffice365Group}>
        {isOffice365Group && <code>{siteService.spfxContext.pageContext.site.group.id.toString()}</code>}
      </DisplayField>

      {/* Support Link */}
      <DisplayField label={strings.SupportLabel} hidden={!supportLink}>
        <a target="_blank" href={supportLink}>{strings.SupportLinkLabel}</a>
      </DisplayField>
    </Stack>
  );
};

export default FooterPanelInner;

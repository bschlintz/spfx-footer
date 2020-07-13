import * as React from 'react';
import { Panel } from 'office-ui-fabric-react/lib/Panel';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import * as strings from 'FooterApplicationCustomizerStrings';
import { SiteService } from '../../../../services/SiteService';
import SponsorField from '../SponsorField/SponsorField';
import SiteAdmins from '../SiteAdmins/SiteAdmins';
import SiteStats from '../SiteStats/SiteStats';
import WebStats from '../WebStats/WebStats';
import DisplayField from '../DisplayField/DisplayField';
import SPPermission from '@microsoft/sp-page-context/lib/SPPermission';

export interface IFooterPanelProps {
  isOpen: boolean;
  onCancelOrDismiss: () => void;
  siteService: SiteService;
  supportLink: string;
}

const FooterPanel: React.FC<IFooterPanelProps> = ({ isOpen, onCancelOrDismiss, siteService, supportLink }) => {
  const isCurrentUserAnAdmin = siteService.spfxContext.pageContext.web.permissions.hasPermission(SPPermission.manageWeb as any);

  return (
    <Panel
      isOpen={isOpen}
      isBlocking={false}
      isLightDismiss={true}
      onDismiss={onCancelOrDismiss}
      headerText={strings.FooterPanelHeaderText}
    >
      <Stack tokens={{ childrenGap: 18 }} style={{ flex: 1 }}>
        {isOpen && <>
          <SponsorField siteService={siteService} />
          <SiteAdmins siteService={siteService} />
          {isCurrentUserAnAdmin && <SiteStats siteService={siteService} />}
          <WebStats siteService={siteService} />
          <DisplayField label={strings.SupportLabel} hidden={!supportLink}>
            <a target="_blank" href={supportLink}>{strings.SupportLinkLabel}</a>
          </DisplayField>
        </>}
      </Stack>
    </Panel>
  );
};

export default FooterPanel;

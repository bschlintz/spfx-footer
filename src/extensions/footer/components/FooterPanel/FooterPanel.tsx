import * as React from 'react';
import * as strings from 'FooterApplicationCustomizerStrings';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { SiteService } from '../../../../services/SiteService';
import FooterPanelInner from './FooterPanelInner';

export interface IFooterPanelProps {
  isOpen: boolean;
  onCancelOrDismiss: () => void;
  siteService: SiteService;
  supportLink: string;
  disableSiteSponsor: boolean;
}

const FooterPanel: React.FC<IFooterPanelProps> = ({ isOpen, onCancelOrDismiss, siteService, supportLink, disableSiteSponsor }) => {

  return (
    <Panel
      isOpen={isOpen}
      isBlocking={false}
      isLightDismiss={true}
      onDismiss={onCancelOrDismiss}
      headerText={strings.FooterPanelHeaderText}
      type={PanelType.custom}
      customWidth="380px"
    >
      {isOpen && <FooterPanelInner siteService={siteService} supportLink={supportLink} disableSiteSponsor={disableSiteSponsor} />}
    </Panel>
  );
};

export default FooterPanel;

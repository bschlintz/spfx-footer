import * as React from 'react';
import { useState } from 'react';
import styles from './Footer.module.scss';
import FooterPanel from '../FooterPanel/FooterPanel';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { SiteService } from '../../../../services/SiteService';
import * as strings from 'FooterApplicationCustomizerStrings';

export interface IFooterProps {
  siteService: SiteService;
  copyrightText: string;
  supportLink: string;
}

const Footer: React.FC<IFooterProps> = ({ siteService, copyrightText, supportLink }) => {
  const [isPanelOpen, setPanelOpen] = useState<boolean>(false);

  const onSiteSummaryClick = () => {
    setPanelOpen(true);
  };

  const closePanel = () => {
    setPanelOpen(false);
  };

  return (
    <>
      <div className={styles.footerContainer}>
        <div> {/* LEFT */}
          <div onClick={onSiteSummaryClick} className={styles.siteSummaryLinkContainer}>
            <Icon iconName="ContactList" className={styles.siteSummaryIcon} />
            <span className={styles.siteSummaryLinkText}>{strings.FooterOpenPanelText}</span>
          </div>
        </div>

        <div> {/* CENTER */}
          <div className={styles.siteCopyright}>{copyrightText}</div>
        </div>

        <div> {/* RIGHT */}
        </div>
      </div>
      <FooterPanel isOpen={isPanelOpen} onCancelOrDismiss={closePanel} siteService={siteService} supportLink={supportLink} />
    </>
  );
};

export default Footer;

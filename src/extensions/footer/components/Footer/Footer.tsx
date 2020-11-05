import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './Footer.module.scss';
import FooterPanel from '../FooterPanel/FooterPanel';
import { Log } from '@microsoft/sp-core-library';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { SiteService } from '../../../../services/SiteService';
import * as strings from 'FooterApplicationCustomizerStrings';
import { LOG_SOURCE } from '../../FooterApplicationCustomizer';
import useColorStyle from '../../../../hooks/useColorStyle';

export interface IFooterProps {
  siteService: SiteService;
  copyrightText: string;
  supportLink: string;
  disableSiteSponsor: boolean;
  footerBackgroundColor: string;
  footerForegroundColor: string;
}

const Footer: React.FC<IFooterProps> = ({ siteService, copyrightText, supportLink, disableSiteSponsor, footerBackgroundColor, footerForegroundColor }) => {
  const [isPanelOpen, setPanelOpen] = useState<boolean>(false);
  const backgroundColorStyle = useColorStyle(footerBackgroundColor, 'backgroundColor');
  const foregroundColorStyle = useColorStyle(footerForegroundColor, 'color');

  useEffect(() => {
    if (!!footerBackgroundColor && !backgroundColorStyle) {
      Log.warn(LOG_SOURCE, `Invalid color value of '${footerBackgroundColor}' found for extension property FooterBackgroundColor. Falling back to theme colors.`);
    }
    if (!!footerForegroundColor && !foregroundColorStyle) {
      Log.warn(LOG_SOURCE, `Invalid color value of '${footerForegroundColor}' found for extension property FooterForegroundColor. Falling back to theme colors.`);
    }
  }, [footerBackgroundColor, footerForegroundColor, backgroundColorStyle, foregroundColorStyle]);

  const onSiteSummaryClick = () => {
    setPanelOpen(true);
  };

  const closePanel = () => {
    setPanelOpen(false);
  };

  return (
    <>
      <div className={styles.footerContainer} style={{...backgroundColorStyle, ...foregroundColorStyle}}>
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
      <FooterPanel isOpen={isPanelOpen} onCancelOrDismiss={closePanel} siteService={siteService} supportLink={supportLink} disableSiteSponsor={disableSiteSponsor} />
    </>
  );
};

export default Footer;

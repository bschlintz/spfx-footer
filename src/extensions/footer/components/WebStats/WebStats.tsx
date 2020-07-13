import * as React from 'react';
import { SiteService } from '../../../../services/SiteService';
import useAsyncData from '../../../../hooks/useAsyncData';
import * as strings from 'FooterApplicationCustomizerStrings';
import DisplayField from '../DisplayField/DisplayField';
import { getWebTemplateFriendlyName } from '../../../../services/Utils';
import { WebStats } from '../../../../models/WebStats';

export interface ISiteAdminsProps {
  siteService: SiteService;
}

const WebStats: React.FC<ISiteAdminsProps> = ({ siteService }) => {
  const {
    data: webStats,
    isLoading: isLoadingWebStats
  } = useAsyncData<WebStats>(null, siteService.getWebStats, []);

  return (
    <>
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
    </>
  );
};

export default WebStats;

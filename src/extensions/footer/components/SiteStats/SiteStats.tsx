import * as React from 'react';
import { SiteService } from '../../../../services/SiteService';
import useAsyncData from '../../../../hooks/useAsyncData';
import * as strings from 'FooterApplicationCustomizerStrings';
import { SiteStats } from '../../../../models/SiteStats';
import DisplayField from '../DisplayField/DisplayField';

export interface ISiteAdminsProps {
  siteService: SiteService;
}

const SiteStats: React.FC<ISiteAdminsProps> = ({ siteService }) => {
  const {
    data: siteStats,
    isLoading: isLoadingSiteStats
  } = useAsyncData<SiteStats>(null, siteService.getSiteStats, []);

  return (
    <>
      {/* Storage */}
      <DisplayField label={strings.StorageLabel} isLoading={isLoadingSiteStats}>
        {siteStats && siteStats.storageUsedBytes > 0
          ? <span>{(siteStats.storageUsedBytes / 1024 / 1024).toFixed(2)} MB used, {Math.floor(100 - siteStats.storageUsedPercentage)}% free</span>
          : <span>{strings.NoDataFieldLabel}</span>
        }
      </DisplayField>
    </>
  );
};

export default SiteStats;

import * as React from 'react';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { SiteService } from '../../../../services/SiteService';
import useAsyncData from '../../../../hooks/useAsyncData';
import * as strings from 'FooterApplicationCustomizerStrings';
import DisplayField from '../DisplayField/DisplayField';
import Person from '../Person/Person';
import { SiteAdmins } from '../../../../models/SiteAdmins';

export interface ISiteAdminsProps {
  siteService: SiteService;
}

const SiteAdmins: React.FC<ISiteAdminsProps> = ({ siteService }) => {

  const {
    data: siteAdmins,
    isLoading: isLoadingSiteAdmins
  } = useAsyncData<SiteAdmins>(null, siteService.getSiteAdmins, []);

  return (
    <>
      {/* Primary Site Administrator  */}
      <DisplayField label={strings.PrimarySiteAdminLabel} isLoading={isLoadingSiteAdmins}>
        {siteAdmins && siteAdmins.primaryAdmin
          ? <Person loginNameOrUpn={siteAdmins.primaryAdmin.loginName} />
          : <span>{strings.EmptyFieldLabel}</span>
        }
      </DisplayField>

      {/* Site Administrators */}
      <DisplayField label={strings.SiteAdminsLabel} hidden={isLoadingSiteAdmins || !siteAdmins || (siteAdmins && siteAdmins.allAdmins && siteAdmins.allAdmins.length === 0)}>
        <Stack tokens={{ childrenGap: 6 }}>
          {siteAdmins && siteAdmins.allAdmins && siteAdmins.allAdmins.map(sa => <Person loginNameOrUpn={sa.loginName} />)}
        </Stack>
      </DisplayField>
    </>
  );
};

export default SiteAdmins;

import * as React from 'react';
import { useState } from 'react';
import styles from './SponsorField.module.scss';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { SiteService } from '../../../../services/SiteService';
import useAsyncData from '../../../../hooks/useAsyncData';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import * as strings from 'FooterApplicationCustomizerStrings';
import { SiteUser } from '../../../../models/SiteUser';
import Person from '../Person/Person';

export interface ISponsorField {
  siteService: SiteService;
}

enum FieldMode {
  Loading,
  ReadOnly,
  CanEdit,
  Editing
}

const SponsorField: React.FC<ISponsorField> = ({ siteService }) => {

  const [ fieldMode, setFieldMode ] = useState<FieldMode>(FieldMode.Loading);
  const [ peoplePickerItems, setPeoplePickerItems ] = useState([]);

  const fetchSiteSponsor = async (): Promise<SiteUser> => {
    setFieldMode(FieldMode.Loading);

    const [ siteSponsorResult, userRightsResult ] = await Promise.all([
      siteService.getSiteSponsor(),
      siteService.getUserRights()
    ]);

    setFieldMode(userRightsResult.isSiteSponsorEditor ? FieldMode.CanEdit : FieldMode.ReadOnly);
    return siteSponsorResult;
  };
  const {
    data: siteSponsor,
    replace: replaceSiteSponsor,
  } = useAsyncData<SiteUser>(null, fetchSiteSponsor, []);

  const onCancelClick = () => {
    setFieldMode(FieldMode.CanEdit);
  };

  const onEditClick = () => {
    setFieldMode(FieldMode.Editing);
  };

  const onSaveClick = async () => {
    let siteSponsorLoginName = "";

    if (peoplePickerItems.length > 0) {
      const item = peoplePickerItems[0];
      siteSponsorLoginName = item.loginName;
    }

    setFieldMode(FieldMode.Loading);

    const newSiteSponsor = await siteService.setSiteSponsor(siteSponsorLoginName);
    replaceSiteSponsor(newSiteSponsor);

    setFieldMode(FieldMode.CanEdit);
  };

  return (
    <Stack>
      <Stack horizontal verticalAlign="center" horizontalAlign="space-between" className={styles.fieldControlHeader}>
        <strong>{strings.SiteSponsorLabel}</strong>
        {fieldMode === FieldMode.CanEdit && (
          <Stack>
            <IconButton iconProps={{iconName: "Edit"}} onClick={onEditClick} />
          </Stack>
        )}
      </Stack>

      {fieldMode === FieldMode.Loading && (
        <Spinner size={SpinnerSize.medium} />
      )}

      {(fieldMode === FieldMode.ReadOnly || fieldMode === FieldMode.CanEdit) && <>
        {siteSponsor
          ? <Person user={siteSponsor} />
          : <span>{strings.EmptyFieldLabel}</span>
        }
      </>}

      {fieldMode === FieldMode.Editing && (
        <Stack>
          <PeoplePicker
            context={siteService.spfxContext}
            personSelectionLimit={1}
            placeholder={strings.PeoplePickerPlaceholderText}
            principalTypes={[PrincipalType.User]}
            resolveDelay={500}
            selectedItems={setPeoplePickerItems}
            defaultSelectedUsers={siteSponsor ? [siteSponsor.email] : []}
          />
          <Stack horizontalAlign="end" horizontal>
            <IconButton iconProps={{iconName: "Cancel"}} onClick={onCancelClick} />
            <IconButton iconProps={{iconName: "Accept"}} onClick={onSaveClick} />
          </Stack>
        </Stack>
      )}
    </Stack>
  );
};

export default SponsorField;

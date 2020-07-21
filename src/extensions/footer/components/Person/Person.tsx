import * as React from 'react';
import { toUpn } from '../../../../services/Utils';
import { SiteUser } from '../../../../models/SiteUser';
import { Persona } from 'office-ui-fabric-react/lib/Persona';
import { PrincipalType } from '@pnp/sp';

declare global {
  namespace JSX {
    interface IntrinsicElements {
      'mgt-person': any;
    }
  }
}

export interface IPersonProps {
  user: SiteUser;
}

const Person: React.FC<IPersonProps> = ({ user: { loginName, title, principalType } }) => {

  const renderNonUser = () => {
    let groupLabel = "Group";
    if (principalType === PrincipalType.SecurityGroup) groupLabel = "Security Group";
    if (principalType === PrincipalType.DistributionList) groupLabel = "Distribution List";

    return (
      <Persona text={title} secondaryText={groupLabel} />
    );
  };

  return (
    <div style={{ minHeight: 48 }}>
      {principalType === PrincipalType.User
        ? <mgt-person person-query={toUpn(loginName)} view="twoLines" person-card="hover" show-presence="true"></mgt-person>
        : renderNonUser()
      }
    </div>
  );
};

export default Person;

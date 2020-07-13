import * as React from 'react';
import { toUpn } from '../../../../services/Utils';

declare global {
  namespace JSX {
    interface IntrinsicElements {
      'mgt-person': any;
    }
  }
}

export interface IPersonProps {
  loginNameOrUpn: string;
}

const Person: React.FC<IPersonProps> = ({ loginNameOrUpn }) => {
  return (
    <div style={{ minHeight: 48 }}>
      <mgt-person person-query={toUpn(loginNameOrUpn)} view="twoLines" person-card="hover" show-presence="true"></mgt-person>
    </div>
  );
};

export default Person;

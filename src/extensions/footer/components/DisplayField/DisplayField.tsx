import * as React from 'react';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';

export interface IDisplayFieldProps {
  label: string;
  isLoading?: boolean;
  hidden?: boolean;
}

const DisplayField: React.FC<IDisplayFieldProps> = ({ label, isLoading, hidden, children }) => (
  <>
    {!hidden &&
      <Stack>
        <Stack verticalAlign="center" style={{ minHeight: 32 }}>
          <strong>{label}</strong>
        </Stack>
        {isLoading
          ? <Spinner size={SpinnerSize.medium} />
          : <>{children}</>
        }
      </Stack>
    }
  </>
);

export default DisplayField;

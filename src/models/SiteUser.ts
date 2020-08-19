import { PrincipalType } from "@pnp/sp";

export type SiteUser = {
  title: string;
  loginName: string;
  email: string;
  principalType: PrincipalType | 'unknown';
};

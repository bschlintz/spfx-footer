import { SiteUser } from "./SiteUser";

export type SiteAdmins = {
  primaryAdmin?: SiteUser,
  allAdmins: SiteUser[]
};

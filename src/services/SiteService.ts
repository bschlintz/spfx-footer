
import { ExtensionContext } from '@microsoft/sp-extension-base';
import { Log } from '@microsoft/sp-core-library';
import { LOG_SOURCE } from '../extensions/footer/FooterApplicationCustomizer';
import { SiteUser } from '../models/SiteUser';
import { UserRights } from '../models/UserRights';
import { ConfigListItem } from '../models/ConfigListItem';

import { setup as pnpSetup } from "@pnp/common";
import { sp } from "@pnp/sp";
import { IWebEnsureUserResult } from "@pnp/sp/site-users/";
import "@pnp/sp/site-users/";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { graph } from "@pnp/graph";
import "@pnp/graph/users";
import "@pnp/graph/groups";
import { SiteStats } from '../models/SiteStats';
import { SiteAdmins } from '../models/SiteAdmins';
import { WebStats } from '../models/WebStats';

const CONFIG_LIST_TITLE = "SiteConfig";
const SITE_SPONSOR_ITEM_TITLE = "SITE_SPONSOR";
const SITE_PRIMARY_ADMIN_ITEM_TITLE = "SITE_PRIMARY_ADMIN";

export class SiteService {
  private _context: ExtensionContext;
  private _siteSponsorsAadGroupId: string;

  // In-memory cache
  private _cachedSiteSponsorConfigItem: ConfigListItem = null;
  private _cachedUserRights: UserRights = null;

  constructor(context: ExtensionContext, siteSponsorsAadGroupId: string) {
    this._context = context;
    this._siteSponsorsAadGroupId = siteSponsorsAadGroupId;
    pnpSetup({ spfxContext: context });
  }

  get spfxContext(): ExtensionContext {
    return this._context || null;
  }

  private _getConfigListItem = async (title: string): Promise<ConfigListItem> => {
    try {
      const result = await sp.web.lists.getByTitle(CONFIG_LIST_TITLE)
                                 .items.filter(`Title eq '${title}'`)
                                 .select('ID', 'Title', 'Value').top(1).get();

      return result && result.length > 0 ? result[0] : null;
    }
    catch (error) {
      throw new Error(`Unable to retrieve '${title}' list item from '${CONFIG_LIST_TITLE}' list. Error: ${error}`);
    }
  }

  private _setConfigListItem = async (id: number, value: any): Promise<void> => {
    try {
      await sp.web.lists.getByTitle(CONFIG_LIST_TITLE)
                        .items.getById(id)
                        .update({ Value: value });
    }
    catch (error) {
      throw new Error(`Unable to update item ID '${id}' list item from '${CONFIG_LIST_TITLE}' list. Error: ${error}`);
    }
  }

  private _getIsMemberOfGroup = async (groupId: string): Promise<boolean> => {
    try {
      const result = await graph.me.checkMemberGroups([groupId]);
      const isMember = result.indexOf(groupId) !== -1;
      return isMember;
    }
    catch (error) {
      throw new Error(`Unable to check group membership for Group ID '${groupId}'. Error: ${error}`);
    }
  }

  public getSiteSponsor = async (): Promise<SiteUser> => {
    let siteSponsor: SiteUser = null;
    try {
      this._cachedSiteSponsorConfigItem = await this._getConfigListItem(SITE_SPONSOR_ITEM_TITLE);

      if (this._cachedSiteSponsorConfigItem) {
        const value = this._cachedSiteSponsorConfigItem.Value;
        if (value && value.trim() !== "") {
          const siteSponsorUser: IWebEnsureUserResult = await sp.web.ensureUser(value);
          siteSponsor = {
            loginName: siteSponsorUser.data.LoginName,
            email: siteSponsorUser.data.Email,
            title: siteSponsorUser.data.Title
          };
          Log.info(LOG_SOURCE, `${SITE_SPONSOR_ITEM_TITLE} Ensured User: ${value}`);
        }
      }
    }
    catch (error) {
      Log.error(LOG_SOURCE, error);
    }
    finally {
      return siteSponsor;
    }
  }

  public setSiteSponsor = async (newSiteSponsorLoginName: string): Promise<SiteUser> => {
    try {
      let siteSponsorItem = this._cachedSiteSponsorConfigItem;
      if (!siteSponsorItem) {
        siteSponsorItem = await this._getConfigListItem(SITE_SPONSOR_ITEM_TITLE);
      }

      await this._setConfigListItem(siteSponsorItem.ID, newSiteSponsorLoginName);
      return await this.getSiteSponsor();
    }
    catch (error) {
      Log.error(LOG_SOURCE, error);
    }
  }

  public getUserRights = async (): Promise<UserRights> => {
    let userRights: UserRights = this._cachedUserRights;
    try {
      if (!userRights) {
        const isSiteSponsorEditor = await this._getIsMemberOfGroup(this._siteSponsorsAadGroupId);
        this._cachedUserRights = userRights = { isSiteSponsorEditor };
      }
    }
    catch (error) {
      Log.error(LOG_SOURCE, error);
    }
    finally {
      return userRights;
    }
  }

  public getSiteAdmins = async (): Promise<SiteAdmins> => {
    try {
      const [ primarySiteAdminItem, allSiteAdmins ] = await Promise.all([
        this._getConfigListItem(SITE_PRIMARY_ADMIN_ITEM_TITLE),
        sp.web.siteUsers.filter(`IsSiteAdmin eq true`).select('Title', 'Email', 'LoginName').top(4).get()
      ]);

      let primaryAdmin: SiteUser = null;
      if (primarySiteAdminItem) {
        const value = primarySiteAdminItem.Value;
        if (value && value.trim() !== "") {
          const primaryAdminUser: IWebEnsureUserResult = await sp.web.ensureUser(value);
          primaryAdmin = {
            loginName: primaryAdminUser.data.LoginName,
            email: primaryAdminUser.data.Email,
            title: primaryAdminUser.data.Title
          };
        }
      }

      return {
        primaryAdmin,
        allAdmins: allSiteAdmins.map(sa => ({
          title: sa.Title,
          email: sa.Email,
          loginName: sa.LoginName
        }))
      };
    }
    catch (error){
      Log.error(LOG_SOURCE, error);
    }
  }

  public getSiteStats = async (): Promise<SiteStats> => {
    try {
      const site = await sp.site.select('Usage').get();

      return {
        storageUsedBytes: site.Usage.Storage,
        storageUsedPercentage: site.Usage.StoragePercentageUsed,
      };
    }
    catch (error){
      Log.error(LOG_SOURCE, error);
      return null;
    }
  }

  public getWebStats = async (): Promise<WebStats> => {
    try {
      const web = await sp.web.select('Created', 'LastItemModifiedDate', 'Configuration', 'WebTemplate').get();

      return {
        created: new Date(web.Created),
        lastUpdated: new Date(web.LastItemModifiedDate),
        webTemplate: `${web.WebTemplate}#${web.Configuration}`,
      };
    }
    catch (error){
      Log.error(LOG_SOURCE, error);
      return null;
    }
  }

}


import { ExtensionContext } from '@microsoft/sp-extension-base';
import { Log } from '@microsoft/sp-core-library';
import { LOG_SOURCE } from '../extensions/footer/FooterApplicationCustomizer';
import { SiteUser } from '../models/SiteUser';
import { UserRights } from '../models/UserRights';
import { ConfigListItem } from '../models/ConfigListItem';
import { AadHttpClient } from '@microsoft/sp-http';
import { setup as pnpSetup } from "@pnp/common";
import { sp, PrincipalType } from "@pnp/sp";
import { IWebEnsureUserResult } from "@pnp/sp/site-users/";
import "@pnp/sp/site-users/";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { SiteStats } from '../models/SiteStats';
import { WebStats } from '../models/WebStats';
import { toUpn } from './Utils';
const CONFIG_LIST_TITLE = "SiteConfig";
const SITE_SPONSOR_ITEM_TITLE = "SITE_SPONSOR";
const SITE_PRIMARY_ADMIN_ITEM_TITLE = "SITE_PRIMARY_ADMIN";

export class SiteService {
  private _context: ExtensionContext;
  private _siteSponsorsAadGroupId: string;
  private _isOffice365Group: boolean = false;
  private _office365GroupId: string;
  private _graphAadClient: AadHttpClient = null;

  // In-memory cache
  private _cachedSiteSponsorConfigItem: ConfigListItem = null;
  private _cachedUserRights: UserRights = null;

  constructor(context: ExtensionContext, siteSponsorsAadGroupId: string) {
    this._context = context;
    this._siteSponsorsAadGroupId = siteSponsorsAadGroupId;
    this._isOffice365Group = !!this._context.pageContext.site.group;
    if (this._isOffice365Group) {
      this._office365GroupId = this._context.pageContext.site.group.id.toString();
    }
    pnpSetup({ spfxContext: context });
  }

  get spfxContext(): ExtensionContext {
    return this._context || null;
  }

  private _getGraphAadClient = async (): Promise<AadHttpClient> => {
    if (!this._graphAadClient) {
      this._graphAadClient = await this._context.aadHttpClientFactory.getClient('https://graph.microsoft.com');
    }
    return this._graphAadClient;
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
      const client = await this._getGraphAadClient();
      let headers: Headers = new Headers();
      headers.set('Accept', 'application/json');
      headers.set('Content-Type', 'application/json');

      const body = JSON.stringify({ groupIds: [ groupId ] });
      const response = await client.post(
        `https://graph.microsoft.com/v1.0/me/checkMemberGroups`,
         AadHttpClient.configurations.v1,
         { headers, body }
      );

      if (response.ok) {
        const result = await response.json();
        const isMember = result.value.indexOf(groupId) !== -1;
        return isMember;
      }
      else throw new Error(await response.text());
    }
    catch (error) {
      throw new Error(`Unable to check group membership for Group ID '${groupId}'. Error: ${error}`);
    }
  }

  private _getGroupOwners = async (top: number = 4): Promise<SiteUser[]> => {
    try {
      const client = await this._getGraphAadClient();
      let headers: Headers = new Headers();
      headers.set('Accept', 'application/json');

      const response = await client.get(
        `https://graph.microsoft.com/v1.0/groups/${this._office365GroupId}/owners?$select=mail,userPrincipalName,displayName&$top=${top}`,
        AadHttpClient.configurations.v1,
        { headers }
      );

      if (response.ok) {
        const result = await response.json();
        return (result.value as any[]).map<SiteUser>(owner => {
          const principalType = owner['@odata.type'] !== '#microsoft.graph.user' ? PrincipalType.SecurityGroup : PrincipalType.User;
          return this._makeSiteUser(owner.userPrincipalName, owner.mail, owner.displayName, principalType);
        });
      }
      else throw new Error(await response.text());
    }
    catch (error) {
      throw new Error(`Unable to fetch group owners for Group ID '${this._office365GroupId}'. Error: ${error}`);
    }
  }

  private _getPerson = async (siteUser: SiteUser): Promise<any> => {
    try {
      const client = await this._getGraphAadClient();
      let headers: Headers = new Headers();
      headers.set('Accept', 'application/json');
      headers.set('Content-Type', 'application/json');
      // Find users without EXO Mailboxes
      // https://docs.microsoft.com/en-us/graph/people-example#types-of-results-included
      headers.set('X-PeopleQuery-QuerySources', 'Mailbox,Directory');

      const response = await client.get(
        `https://graph.microsoft.com/v1.0/me/people?$search="${toUpn(siteUser.loginName)}"&$top=1&$filter=personType/class eq 'Person'`,
         AadHttpClient.configurations.v1,
         { headers }
      );

      if (response.ok) {
        const result = await response.json();
        let person = null;
        if (result.value && result.value.length > 0) {
          person = result.value[0];
        }
        else {
          person = {
            displayName: siteUser.title || siteUser.loginName,
            mail: siteUser.email,
          };
        }
        return person;
      }
      else throw new Error(await response.text());
    }
    catch (error) {
      throw new Error(`Unable to retrieve person with Login Name '${siteUser.loginName}'. Error: ${error}`);
    }
  }

  private _getSiteAdmins = async (top: number = 4): Promise<SiteUser[]> => {
    const siteAdmins = await sp.web.siteUsers.filter(`IsSiteAdmin eq true`).select('Title', 'Email', 'LoginName', 'PrincipalType').top(top).get();
    return siteAdmins.map(sa => this._makeSiteUser(sa.LoginName, sa.Email, sa.Title, sa.PrincipalType));
  }

  private _makeSiteUser = (loginName: string, email: string = '', title: string = '', principalType: PrincipalType | 'unknown' = 'unknown'): SiteUser => {
    return {
      loginName,
      email,
      title,
      principalType
    };
  }

  public getPersonDetails = async (siteUser: SiteUser): Promise<any> => {
    let person = null;
    try {
      person = this._getPerson(siteUser);
    }
    catch (error){
      Log.error(LOG_SOURCE, error);
    }
    finally {
      Log.info(LOG_SOURCE, `getPersonDetails() -> ${JSON.stringify(person)}`);
      return person;
    }
  }

  public getSiteSponsor = async (): Promise<SiteUser> => {
    let siteSponsor: SiteUser = null;
    try {
      this._cachedSiteSponsorConfigItem = await this._getConfigListItem(SITE_SPONSOR_ITEM_TITLE);

      if (this._cachedSiteSponsorConfigItem) {
        const value = this._cachedSiteSponsorConfigItem.Value;
        if (value && value.trim() !== "") {
          let siteSponsorUser: IWebEnsureUserResult = null;
          try {
            Log.info(LOG_SOURCE, `[${SITE_SPONSOR_ITEM_TITLE}] Ensuring user '${value}'`);
            siteSponsorUser = await sp.web.ensureUser(value);
            const { LoginName, Email, Title, PrincipalType } = siteSponsorUser.data;
            siteSponsor = this._makeSiteUser(LoginName, Email, Title, PrincipalType);
          }
          catch (error) {
            Log.error(LOG_SOURCE, new Error(`[${SITE_SPONSOR_ITEM_TITLE}] Unable to ensure user '${value}'`));
            Log.error(LOG_SOURCE, error);
            siteSponsor = this._makeSiteUser(value);
          }
        }
      }
    }
    catch (error) {
      Log.error(LOG_SOURCE, error);
    }
    finally {
      Log.info(LOG_SOURCE, `getSiteSponsor() -> ${JSON.stringify(siteSponsor)}`);
      return siteSponsor;
    }
  }

  public setSiteSponsor = async (newSiteSponsorLoginName: string): Promise<SiteUser> => {
    let newSiteSponsor: SiteUser = null;
    try {
      let siteSponsorItem = this._cachedSiteSponsorConfigItem;
      if (!siteSponsorItem) {
        siteSponsorItem = await this._getConfigListItem(SITE_SPONSOR_ITEM_TITLE);
      }

      await this._setConfigListItem(siteSponsorItem.ID, newSiteSponsorLoginName);
      newSiteSponsor = await this.getSiteSponsor();
    }
    catch (error) {
      Log.error(LOG_SOURCE, error);
    }
    finally {
      Log.info(LOG_SOURCE, `setSiteSponsor(${newSiteSponsorLoginName}) -> ${JSON.stringify(newSiteSponsor)}`);
      return newSiteSponsor;
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
      Log.info(LOG_SOURCE, `getUserRights() -> ${JSON.stringify(userRights)}`);
      return userRights;
    }
  }

  public getPrimaryAdmin = async (): Promise<SiteUser> => {
    let primaryAdmin: SiteUser = null;
    try {
      const primarySiteAdminItem = await this._getConfigListItem(SITE_PRIMARY_ADMIN_ITEM_TITLE);

      if (primarySiteAdminItem) {
        const value = primarySiteAdminItem.Value;
        if (value && value.trim() !== "") {
          let primaryAdminUser: IWebEnsureUserResult = null;
          try {
            Log.info(LOG_SOURCE, `[${SITE_PRIMARY_ADMIN_ITEM_TITLE}] Ensuring user '${value}'`);
            primaryAdminUser = await sp.web.ensureUser(value);
            const { LoginName, Email, Title, PrincipalType } = primaryAdminUser.data;
            primaryAdmin = this._makeSiteUser(LoginName, Email, Title, PrincipalType);
          }
          catch (error) {
            Log.error(LOG_SOURCE, new Error(`[${SITE_PRIMARY_ADMIN_ITEM_TITLE}] Unable to ensure user '${value}'`));
            Log.error(LOG_SOURCE, error);
            primaryAdmin = this._makeSiteUser(value);
          }
        }
      }

      return primaryAdmin;
    }
    catch (error){
      Log.error(LOG_SOURCE, error);
    }
    finally {
      Log.info(LOG_SOURCE, `getPrimaryAdmin() -> ${JSON.stringify(primaryAdmin)}`);
      return primaryAdmin;
    }
  }

  public getSiteAdminsOrGroupOwners = async (): Promise<SiteUser[]> => {
    let adminsOrOwners: SiteUser[] = [];
    try {
      adminsOrOwners = this._isOffice365Group ? await this._getGroupOwners() : await this._getSiteAdmins();
    }
    catch (error){
      Log.error(LOG_SOURCE, error);
    }
    finally {
      Log.info(LOG_SOURCE, `getSiteAdminsOrGroupOwners() -> ${JSON.stringify(adminsOrOwners)}`);
      return adminsOrOwners;
    }
  }

  public getSiteStats = async (): Promise<SiteStats> => {
    let siteStats: SiteStats = null;
    try {
      const site = await sp.site.select('Usage').get();

      siteStats = {
        storageUsedBytes: site.Usage.Storage,
        storageUsedPercentage: site.Usage.StoragePercentageUsed,
      };
    }
    catch (error){
      Log.error(LOG_SOURCE, error);
    }
    finally {
      Log.info(LOG_SOURCE, `getSiteStats() -> ${JSON.stringify(siteStats)}`);
      return siteStats;
    }
  }

  public getWebStats = async (): Promise<WebStats> => {
    let webStats: WebStats = null;
    try {
      const web = await sp.web.select('Created', 'LastItemModifiedDate', 'Configuration', 'WebTemplate').get();

      webStats = {
        created: new Date(web.Created),
        lastUpdated: new Date(web.LastItemModifiedDate),
        webTemplate: `${web.WebTemplate}#${web.Configuration}`,
      };
    }
    catch (error){
      Log.error(LOG_SOURCE, error);
      return null;
    }
    finally {
      Log.info(LOG_SOURCE, `getWebStats() -> ${JSON.stringify(webStats)}`);
      return webStats;
    }
  }

}

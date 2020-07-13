import * as React from 'react';
import * as ReactDom from 'react-dom';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';
import Footer, { IFooterProps } from './components/Footer/Footer';
import { Providers, SharePointProvider } from '@microsoft/mgt';

import { SiteService } from '../../services/SiteService';

export const LOG_SOURCE: string = 'Footer';

export interface IFooterApplicationCustomizerProperties {
  SiteSponsorEditorsAADGroupId: string;
  CopyrightText: string;
  SupportLink: string;
}

export default class FooterApplicationCustomizer
  extends BaseApplicationCustomizer<IFooterApplicationCustomizerProperties> {

  private _bottomPlaceholder: PlaceholderContent;
  private _siteService: SiteService;

  @override
  public async onInit(): Promise<void> {

    // Initialize Site Service
    this._siteService = new SiteService(this.context, this.properties.SiteSponsorEditorsAADGroupId);

    // Initialize Microsoft Graph Toolkit
    Providers.globalProvider = new SharePointProvider(this.context);

    // Event handler to re-render banner on each page navigation
    this.context.application.navigatedEvent.add(this, this.onNavigated);
  }

  /**
   * Event handler that fires on every page load
   */
  private async onNavigated(): Promise<void> {
    this.renderFooter();
  }

  private parseText(text: string): string {
    if (!text) return text;

    let output = text;
    output = output.replace('{current_year}', `${new Date().getFullYear()}`);
    return output;
  }

  /**
   * Render footer React component
   */
  private renderFooter(): void {
    if (!this._bottomPlaceholder) {
      this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom);

      if (!this._bottomPlaceholder) {
        Log.error(LOG_SOURCE, new Error(`Unable to render Bottom placeholder`));
        return;
      }
    }

    //Render Banner React component
    const footerProps: IFooterProps = {
      siteService: this._siteService,
      copyrightText: this.parseText(this.properties.CopyrightText),
      supportLink: this.properties.SupportLink
    };

    const footerComponent = React.createElement(Footer, footerProps);

    ReactDom.render(footerComponent, this._bottomPlaceholder.domElement);
  }
}



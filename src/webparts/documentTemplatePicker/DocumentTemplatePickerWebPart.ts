import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { spfi } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/search";
import { SPFx } from "@pnp/sp";

import * as strings from 'DocumentTemplatePickerWebPartStrings';
import DocumentTemplatePicker from './components/DocumentTemplatePicker';
import { IDocumentTemplatePickerProps } from './components/IDocumentTemplatePickerProps';

export interface IDocumentTemplatePickerWebPartProps {
  webPartTitle?: string;
  templatesSiteUrl?: string;
  templatesLibraryId: string;
  templatesLibraryTitle: string;
  templatesLibraryWebUrl?: string;
  destinationSiteUrl?: string;
  destinationLibraryId: string;
  destinationLibraryTitle: string;
  destinationLibraryWebUrl?: string;
  allowCreateAtRoot: boolean;
  showPreviewColumn: boolean;
}

export default class DocumentTemplatePickerWebPart extends BaseClientSideWebPart<IDocumentTemplatePickerWebPartProps> {

  private _siteOptions: IPropertyPaneDropdownOption[] = [];
  private _templatesLibraryOptions: IPropertyPaneDropdownOption[] = [];
  private _destinationLibraryOptions: IPropertyPaneDropdownOption[] = [];
  private _libraryOptionsLoading: boolean = false;
  private _siteOptionsLoading: boolean = false;

  public render(): void {
    const element: React.ReactElement<IDocumentTemplatePickerProps> = React.createElement(
      DocumentTemplatePicker,
      {
        context: this.context,
        templatesLibraryId: this.properties.templatesLibraryId,
        templatesLibraryTitle: this.properties.templatesLibraryTitle,
        templatesLibraryWebUrl: this.properties.templatesLibraryWebUrl,
        destinationLibraryId: this.properties.destinationLibraryId,
        destinationLibraryTitle: this.properties.destinationLibraryTitle,
        destinationLibraryWebUrl: this.properties.destinationLibraryWebUrl,
        allowCreateAtRoot: this.properties.allowCreateAtRoot !== undefined ? this.properties.allowCreateAtRoot : false,
        showPreviewColumn: this.properties.showPreviewColumn !== undefined ? this.properties.showPreviewColumn : true,
        webPartTitle: this.properties.webPartTitle,
        onConfigure: this._onConfigure.bind(this)
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _onConfigure(): void {
    this.context.propertyPane.open();
  }

  protected async onInit(): Promise<void> {
    await this._loadSiteOptions();
    // Load libraries for templates and destination sites (or current site if not set)
    const templatesSiteUrl = this.properties.templatesSiteUrl || this.context.pageContext.web.absoluteUrl;
    const destinationSiteUrl = this.properties.destinationSiteUrl || this.context.pageContext.web.absoluteUrl;
    await Promise.all([
      this._loadLibraryOptions(templatesSiteUrl, true),
      this._loadLibraryOptions(destinationSiteUrl, false)
    ]);
    return Promise.resolve();
  }

  private async _loadSiteOptions(): Promise<void> {
    if (this._siteOptionsLoading) {
      return;
    }
    this._siteOptionsLoading = true;

    try {
      const sp = spfi().using(SPFx(this.context));
      
      // Use SharePoint Search API to find all sites in the tenant
      const searchQuery = {
        Querytext: "contentclass:STS_Site OR contentclass:STS_Web",
        RowLimit: 500,
        SelectProperties: ['Title', 'Path', 'SiteUrl', 'WebUrl'],
        TrimDuplicates: true
      };

      const searchResults = await sp.search(searchQuery);
      
      // Process search results to extract unique sites/webs
      const siteMap = new Map<string, { url: string; title: string }>();
      
      // Add current site first
      const currentWebUrl = this.context.pageContext.web.absoluteUrl;
      siteMap.set(currentWebUrl, {
        url: currentWebUrl,
        title: `${this.context.pageContext.web.title} (Current)`
      });

      (searchResults.PrimarySearchResults || []).forEach((result: any) => {
        const webUrl = result.WebUrl || result.Path || '';
        const siteUrl = result.SiteUrl || '';
        const title = result.Title || '';
        
        if (webUrl && !siteMap.has(webUrl)) {
          // Use WebUrl if available, otherwise use Path
          const url = webUrl.startsWith('http') ? webUrl : (siteUrl || webUrl);
          if (url) {
            siteMap.set(webUrl, {
              url: url,
              title: title || url
            });
          }
        }
      });

      // Convert to dropdown options
      this._siteOptions = Array.from(siteMap.values())
        .sort((a, b) => a.title.localeCompare(b.title))
        .map((site) => ({
          key: site.url,
          text: site.title
        }));
    } catch (error) {
      console.error('Error loading site options:', error);
      // Fallback to current site only
      this._siteOptions = [{
        key: this.context.pageContext.web.absoluteUrl,
        text: `${this.context.pageContext.web.title} (Current)`
      }];
    } finally {
      this._siteOptionsLoading = false;
    }
  }

  private async _loadLibraryOptions(siteUrl?: string, isTemplates: boolean = true): Promise<void> {
    if (this._libraryOptionsLoading) {
      return;
    }
    this._libraryOptionsLoading = true;

    try {
      const targetSiteUrl = siteUrl || this.context.pageContext.web.absoluteUrl;
      const sp = targetSiteUrl !== this.context.pageContext.web.absoluteUrl
        ? spfi(targetSiteUrl).using(SPFx(this.context))
        : spfi().using(SPFx(this.context));

      // Get libraries from the selected site/web
      const lists = await sp.web.lists
        .filter("BaseTemplate eq 101 and Hidden eq false")
        .select('Id', 'Title')
        .orderBy('Title')();
      
      const options: IPropertyPaneDropdownOption[] = lists.map((list: any) => ({
        key: `${list.Id}|${targetSiteUrl}`,
        text: list.Title
      }));

      if (isTemplates) {
        this._templatesLibraryOptions = options;
      } else {
        this._destinationLibraryOptions = options;
      }
    } catch (error) {
      console.error('Error loading library options:', error);
      // Fallback to empty options
      if (isTemplates) {
        this._templatesLibraryOptions = [];
      } else {
        this._destinationLibraryOptions = [];
      }
    } finally {
      this._libraryOptionsLoading = false;
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    console.log('templatesLibraryOptions',this._templatesLibraryOptions);
    console.log('destinationLibraryOptions',this._destinationLibraryOptions);
    console.log('selectedTemplatesLibrary',`${this.properties.templatesLibraryId}|${this.properties.templatesLibraryWebUrl || this.context.pageContext.web.absoluteUrl}`);
    console.log('selectedDestinationLibrary',`${this.properties.destinationLibraryId}|${this.properties.destinationLibraryWebUrl || this.context.pageContext.web.absoluteUrl}`);
    console.log('seletctedT',this._templatesLibraryOptions.find(option => option.key === `${this.properties.templatesLibraryId}|${this.properties.templatesLibraryWebUrl || this.context.pageContext.web.absoluteUrl}`));
    
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: 'Web Part Title',
              groupFields: [
                PropertyPaneTextField('webPartTitle', {
                  label: 'Title',
                  description: 'Enter a custom title to display at the top of the web part',
                  value: this.properties.webPartTitle
                })
              ]
            },
            {
              groupName: strings.TemplatesLibraryGroupName,
              groupFields: [
                PropertyPaneDropdown('templatesSiteUrl', {
                  label: 'Templates Site',
                  options: this._siteOptions,
                  selectedKey: this.properties.templatesSiteUrl || this.context.pageContext.web.absoluteUrl
                }),
                PropertyPaneDropdown('templatesLibraryId', {
                  label: strings.TemplatesLibraryFieldLabel,
                  options: this._templatesLibraryOptions,
                  selectedKey: `${this.properties.templatesLibraryId}|${this.properties.templatesLibraryWebUrl || this.context.pageContext.web.absoluteUrl}`
                }),
                PropertyPaneCheckbox('showPreviewColumn', {
                  text: 'Show Preview column',
                  checked: this.properties.showPreviewColumn !== undefined ? this.properties.showPreviewColumn : true
                })
              ]
            },
            {
              groupName: strings.DestinationLibraryGroupName,
              groupFields: [
                PropertyPaneDropdown('destinationSiteUrl', {
                  label: 'Destination Site',
                  options: this._siteOptions,
                  selectedKey: this.properties.destinationSiteUrl || this.context.pageContext.web.absoluteUrl
                }),
                PropertyPaneDropdown('destinationLibraryId', {
                  label: strings.DestinationLibraryFieldLabel,
                  options: this._destinationLibraryOptions,
                  selectedKey: `${this.properties.destinationLibraryId}|${this.properties.destinationLibraryWebUrl || this.context.pageContext.web.absoluteUrl}`
                }),
                PropertyPaneCheckbox('allowCreateAtRoot', {
                  text: 'Allow creating documents at root of destination library',
                  checked: this.properties.allowCreateAtRoot !== undefined ? this.properties.allowCreateAtRoot : false
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    // Handle site selection changes
    if (propertyPath === 'templatesSiteUrl' && newValue) {
      this.properties.templatesSiteUrl = newValue;
      // Clear library selection when site changes
      this.properties.templatesLibraryId = '';
      this.properties.templatesLibraryTitle = '';
      this.properties.templatesLibraryWebUrl = newValue;
      // Reload libraries for the selected site
      void this._loadLibraryOptions(newValue, true).then(() => {
        this.context.propertyPane.refresh();
        this.render();
      });
      return;
    }

    if (propertyPath === 'destinationSiteUrl' && newValue) {
      this.properties.destinationSiteUrl = newValue;
      // Clear library selection when site changes
      this.properties.destinationLibraryId = '';
      this.properties.destinationLibraryTitle = '';
      this.properties.destinationLibraryWebUrl = newValue;
      // Reload libraries for the selected site
      void this._loadLibraryOptions(newValue, false).then(() => {
        this.context.propertyPane.refresh();
        this.render();
      });
      return;
    }

    // Handle library selection changes
    if ((propertyPath === 'templatesLibraryId' || propertyPath === 'destinationLibraryId') && newValue) {
      // Parse the key to extract library ID and web URL
      // Format: "libraryId|webUrl" or just "libraryId" for backward compatibility
      const parts = newValue.split('|');
      const libraryId = parts[0];
      const webUrl = parts[1] || this.properties.templatesSiteUrl || this.properties.destinationSiteUrl || this.context.pageContext.web.absoluteUrl;

      if (!libraryId) {
        console.error('Invalid library selection format');
        return;
      }

      // Get the list title from the selected web
      const sp = webUrl !== this.context.pageContext.web.absoluteUrl
        ? spfi(webUrl).using(SPFx(this.context))
        : spfi().using(SPFx(this.context));
      
      sp.web.lists.getById(libraryId)
        .select('Title')()
        .then((listData: any) => {
          if (propertyPath === 'templatesLibraryId') {
            this.properties.templatesLibraryId = libraryId;
            this.properties.templatesLibraryTitle = listData.Title;
            this.properties.templatesLibraryWebUrl = webUrl;
            if (!this.properties.templatesSiteUrl) {
              this.properties.templatesSiteUrl = webUrl;
            }
          } else if (propertyPath === 'destinationLibraryId') {
            this.properties.destinationLibraryId = libraryId;
            this.properties.destinationLibraryTitle = listData.Title;
            this.properties.destinationLibraryWebUrl = webUrl;
            if (!this.properties.destinationSiteUrl) {
              this.properties.destinationSiteUrl = webUrl;
            }
          }
          this.context.propertyPane.refresh();
          this.render();
        })
        .catch((error: any) => {
          console.error('Error fetching library title:', error);
        });
    } else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
  }
}

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { spfi } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import { SPFx } from "@pnp/sp";

import * as strings from 'DocumentTemplatePickerWebPartStrings';
import DocumentTemplatePicker from './components/DocumentTemplatePicker';
import { IDocumentTemplatePickerProps } from './components/IDocumentTemplatePickerProps';

export interface IDocumentTemplatePickerWebPartProps {
  templatesLibraryId: string;
  templatesLibraryTitle: string;
  destinationLibraryId: string;
  destinationLibraryTitle: string;
  allowCreateAtRoot: boolean;
}

export default class DocumentTemplatePickerWebPart extends BaseClientSideWebPart<IDocumentTemplatePickerWebPartProps> {

  private _templatesLibraryOptions: IPropertyPaneDropdownOption[] = [];
  private _destinationLibraryOptions: IPropertyPaneDropdownOption[] = [];
  private _libraryOptionsLoading: boolean = false;

  public render(): void {
    const element: React.ReactElement<IDocumentTemplatePickerProps> = React.createElement(
      DocumentTemplatePicker,
      {
        context: this.context,
        templatesLibraryId: this.properties.templatesLibraryId,
        templatesLibraryTitle: this.properties.templatesLibraryTitle,
        destinationLibraryId: this.properties.destinationLibraryId,
        destinationLibraryTitle: this.properties.destinationLibraryTitle,
        allowCreateAtRoot: this.properties.allowCreateAtRoot !== undefined ? this.properties.allowCreateAtRoot : false,
        onConfigure: this._onConfigure.bind(this)
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _onConfigure(): void {
    this.context.propertyPane.open();
  }

  protected async onInit(): Promise<void> {
    await this._loadLibraryOptions();
    return Promise.resolve();
  }

  private async _loadLibraryOptions(): Promise<void> {
    if (this._libraryOptionsLoading) {
      return;
    }
    this._libraryOptionsLoading = true;

    try {
      const sp = spfi().using(SPFx(this.context));
      
      // Get all document libraries (baseTemplate = 101) excluding hidden lists
      const lists = await sp.web.lists
        .filter("BaseTemplate eq 101 and Hidden eq false")
        .select('Id', 'Title')
        .orderBy('Title')();

      const options: IPropertyPaneDropdownOption[] = lists.map((list: any) => ({
        key: list.Id,
        text: list.Title
      }));

      this._templatesLibraryOptions = options;
      this._destinationLibraryOptions = options;
    } catch (error) {
      console.error('Error loading library options:', error);
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
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.TemplatesLibraryGroupName,
              groupFields: [
                PropertyPaneDropdown('templatesLibraryId', {
                  label: strings.TemplatesLibraryFieldLabel,
                  options: this._templatesLibraryOptions,
                  selectedKey: this.properties.templatesLibraryId
                })
              ]
            },
            {
              groupName: strings.DestinationLibraryGroupName,
              groupFields: [
                PropertyPaneDropdown('destinationLibraryId', {
                  label: strings.DestinationLibraryFieldLabel,
                  options: this._destinationLibraryOptions,
                  selectedKey: this.properties.destinationLibraryId
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
    if ((propertyPath === 'templatesLibraryId' || propertyPath === 'destinationLibraryId') && newValue) {
      // Initialize PnP SP with the current context
      const sp = spfi().using(SPFx(this.context));

      // Get the list title when library is selected
      sp.web.lists.getById(newValue)
        .select('Title')()
        .then((listData: any) => {
          if (propertyPath === 'templatesLibraryId') {
            this.properties.templatesLibraryTitle = listData.Title;
          } else if (propertyPath === 'destinationLibraryId') {
            this.properties.destinationLibraryTitle = listData.Title;
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

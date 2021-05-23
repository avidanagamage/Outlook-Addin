import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'OutlookAddinWebPartStrings';
import { IOutlookAddinProps } from './components/IOutlookAddinProps';
import { OutlookAddin } from './components/OutlookAddin';
import { GraphHelper } from '../../services/GraphHelper';
import { IMail } from '../../models/IMail';

export interface IOutlookAddinWebPartProps {
  description: string;
}

export default class OutlookAddinWebPart extends BaseClientSideWebPart<IOutlookAddinWebPartProps> {

  private graphHelper: GraphHelper;
  public render(): void {
    let mail: IMail = null;
    if (this.context.sdks.office) {
      const item = this.context.sdks.office.context.mailbox.item;
      const itemId = this.context.sdks.office.context.mailbox.convertToRestId(item.itemId, 'v2.0');
      if (item !== null) {
        mail = {
          id: itemId,
          attachments: item.attachments
        };
      }
    } else {
      mail = {
        id: '1',
        attachments: [
          {
            id: '1',
            name: 'Sample 1'
          },
          {
            id: '2',
            name: 'Sample 2'
          }
        ]
      };
    }

    const element: React.ReactElement<IOutlookAddinProps> = React.createElement(
      OutlookAddin,
      {
        description: this.properties.description,
        mail,
        graphHelper: this.graphHelper
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public async onInit() {
    await super.onInit();

    this.graphHelper = new GraphHelper(await this.context.msGraphClientFactory.getClient());
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
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

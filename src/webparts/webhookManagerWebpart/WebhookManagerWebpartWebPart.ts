//Bundle our css- these are styles that can't have unique names
require("src/webparts/webhookManagerWebpart/css/webhookManagerWebpart.css")

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IWebPartContext
} from '@microsoft/sp-webpart-base';

import * as strings from 'webhookManagerWebpartStrings';
import WebhookManagerWebpart from './components/WebhookManagerWebpart';
import { IWebhookManagerWebpartProps } from './components/WebhookManagerWebpart';
import { IWebhookManagerWebpartWebPartProps } from './IWebhookManagerWebpartWebPartProps';
import {ISubscriptionService,SubscriptionService} from '../webhookManagerWebpart/services/index';

export default class WebhookManagerWebpartWebPart extends BaseClientSideWebPart<IWebhookManagerWebpartWebPartProps> {

  private _subscriptionService: ISubscriptionService;

  public render(): void {
   
    this._subscriptionService = new SubscriptionService(this.context.spHttpClient,this.context.pageContext.web.absoluteUrl);

    const element: React.ReactElement<IWebhookManagerWebpartProps > = React.createElement(
      WebhookManagerWebpart,
      {
        description: this.properties.description,
        title:this.properties.title,
        subscriptionService: this._subscriptionService
      }
    );
    
    ReactDom.render(element, this.domElement);

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
                PropertyPaneTextField('title',{label:'Title'}),
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

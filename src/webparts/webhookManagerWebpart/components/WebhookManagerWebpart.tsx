import * as React from 'react';
import styles from './WebhookManagerWebpart.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import SubscriptionListing from './SubscriptionListing';
import {ISubscriptionService,SubscriptionService} from '../services/index';

export interface IWebhookManagerWebpartProps {
  description: string;
  title:string;
  subscriptionService:ISubscriptionService;
}

export default class WebhookManagerWebpart extends React.Component<IWebhookManagerWebpartProps, void> {

  constructor(){
    super();
  }

  public render(): React.ReactElement<IWebhookManagerWebpartProps> {

    return (
      <div className={styles.webhookManagerPart}>
          <h2>SharePoint List Webhook Manager</h2>
        {
          <SubscriptionListing subscriptionService={this.props.subscriptionService}  />
         }
      </div>
    );
  }
}

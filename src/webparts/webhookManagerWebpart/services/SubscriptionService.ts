import {ISubscription, IListSubscription, ISubscriptionService} from '../services';
import { SPHttpClient, SPHttpClientResponse,ISPHttpClientOptions } from '@microsoft/sp-http';

export class SubscriptionService implements ISubscriptionService{

	constructor(private spHttpClient:SPHttpClient, private siteUrl: string){

	}

	public getLists = ():Promise<IListSubscription[]>  => {
		
			return new Promise<IListSubscription[]>((resolve: (results:IListSubscription[]) => void,reject: (error:any) => void) : void =>{

				 this.spHttpClient.get(`${this.siteUrl}/_api/web/lists?$expand=RootFolder,Subscriptions`,
					SPHttpClient.configurations.v1,
					{
						headers: {
						'Accept': 'application/json;odata=nometadata',
						'odata-version': ''
						}
					})
					.then((response: SPHttpClientResponse): Promise<{ value: any }> => {
						return response.json();
					})
					.then((response: { value: any }): void => {
						
						//Get List Details
						let lists:Array<IListSubscription> = response.value.map((list) => {

								var subscriptionList:IListSubscription = {
									title:list.Title,
									description: '',
									uniqueId:list.Id,
									subscriptions: list.Subscriptions.map((sub) => { var subscription:ISubscription;  subscription = {resource:sub.resource, notificationUrl: sub.notificationUrl, expirationDateTime: sub.expirationDateTime, clientState: sub.clientState, id: sub.id}; return subscription;} ),
									subscriptionsCount: list.Subscriptions.length,
									url: list.RootFolder.ServerRelativeUrl
								}
								return subscriptionList;

						});

						resolve(lists);

					}, (error: any): void => {
						reject(error);
					});

			});

	}

	public addSubscription = (subscription:ISubscription, uniqueId:string): Promise<ISubscription> => {

		subscription.resource = `${this.siteUrl}/_api/web/lists('${uniqueId}')`;
		
		return new Promise<ISubscription>((resolve: (results:ISubscription) => void,reject: (error:any) => void) : void =>{

				const spOpts: ISPHttpClientOptions = {
					body: JSON.stringify(subscription)
				};

				 this.spHttpClient.post(`${this.siteUrl}/_api/web/lists('${uniqueId}')/subscriptions`, 
					SPHttpClient.configurations.v1,spOpts)
					.then((response: SPHttpClientResponse) => {
						if (response.status == 201){
							resolve(subscription);
						}
						else
						{
							reject(response.statusText);
						}
					}, (error: any): void => {
						reject(error);
					});

			});
	}

	public deleteSubscription = (listUniqueId:string, uniqueId:string): Promise<void> => {
		
		return new Promise<void>((resolve: () => void,reject: (error:any) => void) : void =>{
				
			 const spOpts: ISPHttpClientOptions = {
				 headers: {
					'X-HTTP-Method': 'DELETE'
					}
			 };

			 this.spHttpClient.post(`${this.siteUrl}/_api/web/lists('${listUniqueId}')/subscriptions('${uniqueId}')`,
					SPHttpClient.configurations.v1,spOpts)
					.then((response: SPHttpClientResponse) => {
						if (response.status == 204){
							resolve();
						}
						else
						{
							reject(response.statusText);
						}
					}, (error: any): void => {
						reject(error);
					});

			});
	}

}

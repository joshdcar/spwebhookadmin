import {ISubscription} from './ISubscription';

export interface IListSubscription{

		title:string;
		description:string;
		url:string;
		uniqueId:string;
		subscriptions: Array<ISubscription>;
		subscriptionsCount:number;
}
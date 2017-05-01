import {IListSubscription} from './IListSubscription';
import {ISubscription} from './ISubscription';

export interface ISubscriptionService{
	getLists: () => Promise<IListSubscription[]>;
	addSubscription: (subscription:ISubscription, uniqueId:string) => Promise<ISubscription>;
	deleteSubscription: (listUniqueId:string, uniqueId:string) => Promise<void>;
}

export interface ISubscription{
	id?:string;
	resource:string
	notificationUrl: string;
	expirationDateTime: Date;
	clientState:string;
}
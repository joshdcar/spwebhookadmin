import * as React from 'react';
import {Link, Dialog,DialogType,DialogFooter, Button, ButtonType, IconButton, TextField, Calendar, DatePicker,DayOfWeek, MessageBar, MessageBarType } from "office-ui-fabric-react";
import {IListSubscription, ISubscriptionService, ISubscription} from '../services/index';
import styles from './WebhookManagerWebpart.module.scss';

export interface INewSubscriptionDialogProps{
	showDialog:boolean,
	subscriptionList: IListSubscription,
	subscriptionService:ISubscriptionService;
	hideDialog: (update:boolean, list: IListSubscription, subscription:ISubscription) => void; 
}

export interface INewSubscriptionDialogState{
	notificationUrl:string;
	expirationDate:Date;
	clientState:string;
	error?:any;
	showError:boolean;
}

const DayPickerStrings = {
  months: [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December'
  ],

  shortMonths: [
    'Jan',
    'Feb',
    'Mar',
    'Apr',
    'May',
    'Jun',
    'Jul',
    'Aug',
    'Sep',
    'Oct',
    'Nov',
    'Dec'
  ],

  days: [
    'Sunday',
    'Monday',
    'Tuesday',
    'Wednesday',
    'Thursday',
    'Friday',
    'Saturday'
  ],

  shortDays: [
    'S',
    'M',
    'T',
    'W',
    'T',
    'F',
    'S'
  ],

  goToToday: 'Go to today'
};

export class NewSubscriptionDialog extends React.Component< INewSubscriptionDialogProps, INewSubscriptionDialogState>{

		public constructor(){
			super();

			var defaultDate = new Date();
			defaultDate.setMonth(defaultDate.getMonth() + 5);

			this.state = {
				notificationUrl:'',
				clientState: '',
				expirationDate: defaultDate,
				showError: false
			}

		}

		public render(): JSX.Element { 

			 return (
					 <Dialog
         				isOpen={ this.props.showDialog }
          				type={ DialogType.close }
          				onDismiss={ this._closeNewDialog}
          				title= {this.props.subscriptionList.title + " - New Webhook Subscription"}
          				isBlocking={ true }
          				closeButtonAriaLabel='Close' >
						  	<div>
								   { this.state.error ? 
								   			<MessageBar
												messageBarType={ MessageBarType.error }
												onDismiss={ () => { this.state.showError = false;  } }>
												We're sorry. An error occured when we tried to add our subscription. ${ this.state.error }</MessageBar> : ''
									}

								  <div>
								  		<TextField label='Notification Url' 
								  					placeholder='Enter a valid notification url'  
													validateOnFocusOut 
													onGetErrorMessage= {this._onNotificationUrlValidate}
													onChanged={this._onNotificationUrlChange}
													//onBlur={this._onNotificationUrlChange} 
													autoSave='yes'
													validateOnLoad={false}
													value={this.state.notificationUrl}  required />
								  </div>
								    <div>
								  		<TextField label='Client State' 
								  					placeholder='Enter a client state value required by your service'  
													onChanged={this._onClientStateChange}
													value={this.state.clientState}  required />
								  </div>
								  <div>
								  			<DatePicker firstDayOfWeek={ DayOfWeek.Monday } strings={ DayPickerStrings }  label='Expiration Date' placeholder='Select an expiration date...' value={this.state.expirationDate}  />
								  </div>
							  </div>
							<DialogFooter>
								<Button buttonType={ ButtonType.primary } onClick={ this._addSubscription }>Save</Button>
								<Button onClick={ this._closeNewDialog }>Cancel</Button>
							</DialogFooter>
						</Dialog>
			 );
		}

		private _onNotificationUrlValidate = (value: any) =>{

			if (value === '')
				return "Notification Url is required";

			var expression:RegExp  = /https?:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,4}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)?/gi;
			var valid:boolean = expression.test(value);

			if(!valid)
				return "Notification Url must be a valid url";

			return '';
			
		}

		private _onClientStateChange = (value:any) =>{

			this.setState({
				notificationUrl: value,
				expirationDate: this.state.expirationDate,
				clientState: value,
				showError: false
			})
		}

		private _onNotificationUrlChange = (value:any) =>{

			this.setState({
				notificationUrl: value,
				expirationDate: this.state.expirationDate,
				clientState: this.state.clientState,
				showError: false
			})
		}

		private _addSubscription = () => {

			var subscription:ISubscription = {
				resource:'',
				notificationUrl: this.state.notificationUrl,
				clientState: this.state.clientState,
				expirationDateTime: this.state.expirationDate
			}

			this.props.subscriptionService.addSubscription(subscription,this.props.subscriptionList.uniqueId).then( (subscription:ISubscription) => {
					this.props.hideDialog(true,this.props.subscriptionList,subscription);
			}).catch((error:any) =>{
				this.state.error = error;
				this.state.showError = true;
			});


			
		}

		private _closeNewDialog = () => {
			this.props.hideDialog(false, this.props.subscriptionList, null);
		}


}
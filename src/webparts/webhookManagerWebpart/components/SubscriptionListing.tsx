import * as React from 'react';
import {SPHttpClient} from '@microsoft/sp-http';
import styles from './WebhookManagerWebpart.module.scss';
import {DetailsList, Selection,SelectionMode, buildColumns, IColumn, 
	ColumnActionsMode,CheckboxVisibility, Link, Dialog,DialogType,
	DialogFooter, Button, ButtonType, IconButton, TextField, Calendar, DatePicker,DayOfWeek, Spinner, SpinnerType } from "office-ui-fabric-react";
import {IListSubscription, ISubscriptionService,ISubscription} from '../services/index';
import {NewSubscriptionDialog} from './NewSubscriptionDialog';
import {MessageDialog} from './MessageDialog';
import {SubscriptionsDialog} from './SubscriptionsDialog'

export interface ISubscriptionListingProps{
	 subscriptionService: ISubscriptionService
}

export interface ISubscriptionListingState{
	lists:Array<IListSubscription>;
	columns?: IColumn[];
	selectionDetails?:IListSubscription;
	showSubscriptionsDialog?: boolean;
	showNewSubscriptionDialog?: boolean;
	showSubscriptionListingDialog?:boolean;
	showMessageDialog?:boolean;
	messageDialogTitle?:string;
	messageDialogMessage?:string;
	loading?:boolean;
}


export default class SubscriptionListing extends React.Component<ISubscriptionListingProps, ISubscriptionListingState>{

	constructor(props:ISubscriptionListingProps, state: ISubscriptionListingState){

		super(props);

		let defaultState:ISubscriptionListingState = {
			lists: [],
			columns: this._buildColumns(),
			selectionDetails: {title: '', description:'', uniqueId:'', url:'', subscriptions:[], subscriptionsCount:0},
			showSubscriptionsDialog:false,
			showNewSubscriptionDialog: false,
			loading:true
		};

		this.state = defaultState;

	}

	public componentDidMount(): void {
    	this._loadListData();
  	}

	public render(): JSX.Element { 

		var state = this.state;
		var loadingElement;

		if(this.state.loading){
			loadingElement = <div className={styles.loading} ><Spinner  type={ SpinnerType.large } label='Loading SharePoint Lists...' /></div>
		}
	
		 return (

				<div className={styles.container}>

					<MessageDialog showDialog={this.state.showMessageDialog} 
												message={this.state.messageDialogMessage} 
												title={this.state.messageDialogTitle} 
												hideDialog={this._hideMessageDialog} />
				
					<DetailsList
						items= {this.state.lists}
						setKey='set'
						checkboxVisibility={CheckboxVisibility.hidden}
						columns= {this.state.columns}
						selectionMode={SelectionMode.none}
						selectionPreservedOnEmptyClick={ true }
						onItemInvoked={ this._listItemInvoked }
						onRenderItemColumn={ this._renderItemColumn }
					/>

					{loadingElement}
					  

					<NewSubscriptionDialog subscriptionService={this.props.subscriptionService} 
															showDialog={this.state.showNewSubscriptionDialog}
															 hideDialog={this._hideNewDialog}
															 subscriptionList= {this.state.selectionDetails}  />
					
					<SubscriptionsDialog   subscriptionService={this.props.subscriptionService} 
															removeSubscription= {this._removeSubscription}
															showDialog={this.state.showSubscriptionListingDialog}
															hideDialog={this._hideSubscriptionListingDialog}
															subscriptionList={this.state.selectionDetails}  />

			</div>
		);
	}

	private _hideMessageDialog = () =>{

			this.setState(
			{ 
				lists: this.state.lists,
				columns: this.state.columns,
				selectionDetails:this.state.selectionDetails,
				showNewSubscriptionDialog: false,
				showSubscriptionsDialog: false,
				showMessageDialog: false,
				messageDialogMessage: '',
				messageDialogTitle: ''
			});
	}

	private _hideSubscriptionListingDialog = () => {

				this.setState(
				{ 
					lists: this.state.lists,
					columns: this.state.columns,
					selectionDetails:this.state.selectionDetails,
					showNewSubscriptionDialog: this.state.showNewSubscriptionDialog,
					showSubscriptionListingDialog: false,
					showSubscriptionsDialog: this.state.showSubscriptionsDialog,
					showMessageDialog: this.state.showMessageDialog
				});

	}

	private _removeSubscription = (list:IListSubscription, subscription:ISubscription)=>{

			var lists = this.state.lists;
			var listIndex:number = lists.indexOf(list);
			var subscriptionIndex = lists[listIndex].subscriptions.indexOf(subscription);
			lists[listIndex].subscriptions.slice(subscriptionIndex,1);
			lists[listIndex].subscriptionsCount--;

			this.setState(
			{ 
				lists: lists,
				columns: this.state.columns,
				selectionDetails:this.state.selectionDetails,
				showNewSubscriptionDialog: this.state.showNewSubscriptionDialog,
				showSubscriptionListingDialog: this.state.showSubscriptionListingDialog,
				showSubscriptionsDialog: this.state.showSubscriptionsDialog,
				showMessageDialog: this.state.showMessageDialog,
			});


	}

	private _hideNewDialog = (update:boolean, list: IListSubscription, subscription:ISubscription) => {

			if(update == true){
				var index:number = this.state.lists.indexOf(list);

				if(index != -1){
					this.state.lists[index].subscriptions.push(subscription);
				}

			}

			this.setState(
			{ 
				lists: this.state.lists,
				columns: this.state.columns,
				selectionDetails:this.state.selectionDetails,
				showNewSubscriptionDialog: false,
				showSubscriptionsDialog: false,
				showMessageDialog: update,
				messageDialogTitle: update ? 'New Subscription' : '',
				messageDialogMessage: update ? 'Your new subscription has been successfully added to ' + list.title: ''
			});

	}

	private _showNewDialog = (item:IListSubscription) => {

			this.setState(
			{ 
				lists: this.state.lists,
				columns: this.state.columns,
				selectionDetails:item,
				showNewSubscriptionDialog: true,
				showSubscriptionsDialog: false
			});

	}
	
	private _showSubscriptionsDialog = (item:IListSubscription) => {

		this.setState(
			{ 
				lists: this.state.lists,
				columns: this.state.columns,
				selectionDetails:item,
				showNewSubscriptionDialog: this.state.showNewSubscriptionDialog,
				showSubscriptionListingDialog: true,
				showSubscriptionsDialog: this.state.showSubscriptionsDialog,
				showMessageDialog: this.state.showMessageDialog,
			});

	}
	
	private _closeSubscriptionDialog = () => {	

		this.setState(
			{ 
				lists: this.state.lists,
				columns: this.state.columns,
				selectionDetails:this.state.selectionDetails,
				showNewSubscriptionDialog: false,
				showSubscriptionsDialog: false
			});

	}
	
	private _listItemInvoked = (item:any,  index: number, ev: Event) => {
		
		this._showSubscriptionsDialog(item);

	}

	private _buildColumns = (): Array<IColumn> =>{

		var columns:Array<IColumn> = new Array<IColumn>();

		let titleColumn:IColumn = {
			key:'title',
			name: 'List Title',
			fieldName: 'title',
			minWidth:200
		}

		columns.push(titleColumn);

		let subscriptionColumn:IColumn = {
			key:'subscriptions',
			name: 'Subscriptions',
			fieldName: 'subscriptionsCount',
			minWidth:200
		}

		columns.push(subscriptionColumn);

		return columns;

	}

	private _renderItemColumn = (item, index, column) => {

		let fieldContent = item[column.fieldName];
		let listItem = item as IListSubscription;

		switch(column.key){

			case 'title':
				let url = item["url"];
				return <Link className={styles.titleLink} data-selection-invoke={ true }>{ fieldContent }</Link>      

			case 'subscriptions':
				return <div>
								<Link className={styles.countLink} data-selection-invoke={ true }>{ fieldContent }</Link>      
								 <IconButton
								 	className={styles.newLink}
									onClick={() => this._showNewDialog(listItem)}
									text="Add"
          							icon='CirclePlus'
          							title='Add Subscription'
          							ariaLabel='Add Subscription' />
						</div>; 

		}

	}

	async _loadListData() {
    	
		let lists:Array<IListSubscription> = await this.props.subscriptionService.getLists();

		this.setState({lists:lists,
								loading:false});
	
  	}


}
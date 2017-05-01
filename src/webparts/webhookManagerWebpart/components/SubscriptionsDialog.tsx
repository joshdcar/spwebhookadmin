import * as React from 'react';
import { IListSubscription, ISubscriptionService, ISubscription } from '../services/index';
import {
	Dialog, DialogType, DialogFooter, Button, ButtonType, Spinner, SpinnerType, DetailsList, Selection, SelectionMode, buildColumns, IColumn,
	ColumnActionsMode, CheckboxVisibility, Link,IconButton,Callout, DirectionalHint,MessageBar, MessageBarType
} from "office-ui-fabric-react";
import styles from './WebhookManagerWebpart.module.scss';

export interface ISubscriptionsDialogProps {
	showDialog: boolean,
	subscriptionList: IListSubscription,
	subscriptionService: ISubscriptionService;
	hideDialog: () => void;
	removeSubscription:(list:IListSubscription, subscription:ISubscription) => void;
}

export interface ISubscriptionsDialogState {
	columns?: IColumn[];
	error?: any;
	showError: boolean;
	showDeleteConfirm: Boolean;
	selectedSubscription?: ISubscription;
	deleteTarget?: any;
	subscriptionList: IListSubscription
}


export class SubscriptionsDialog extends React.Component<ISubscriptionsDialogProps, ISubscriptionsDialogState>{

	constructor(props: ISubscriptionsDialogProps, state: ISubscriptionsDialogState) {

		super(props);

		this.state = {
			showError: false,
			showDeleteConfirm:false,
			columns: this._buildColumns(),
			subscriptionList: this.props.subscriptionList //Not the best of pattern but we want to be able to update the list so we're making a copy into state
		}

	}

	public componentDidMount(): void {

			this.setState({
					subscriptionList: this.props.subscriptionList,
					selectedSubscription:this.state.selectedSubscription,
					showError:this.state.showError,
					showDeleteConfirm: this.state.showDeleteConfirm,
					deleteTarget: this.state.deleteTarget
				});	
    	
  	}	


	public render(): JSX.Element {

		var calloutElement;
		var errorElement;

		if(this.state.showDeleteConfirm){
			calloutElement = <Callout
					 		className={styles.subscriptionCallout}
            				targetElement={ this.state.deleteTarget}
            				directionalHint={ DirectionalHint.rightCenter}
            				coverTarget={ true }
            				isBeakVisible={ false }
							gapSpace={ 0 }>
								<div className={styles.subscriptionCalloutHeader}>
										Are you sure you want to delete this subscription?
								</div>
								<div className={styles.subscriptionCalloutButtons}>
										<Button onClick={  this._onConfirmDeleteSubscription }> Confirm </Button>
										<Button onClick={  this._onCancelDelete }> Cancel </Button>
								</div>
							</Callout>;
		}

		if(this.state.showDeleteConfirm){
			errorElement=<div> <MessageBar messageBarType={ MessageBarType.error } >
      																			We're sorry an error occured while trying to delete the subscription. Error: ${this.state.error}</MessageBar></div>;
		}

		return (

			<Dialog
				isOpen={this.props.showDialog}
				type={DialogType.close}
				onDismiss={this._closeNewDialog}
				title={this.props.subscriptionList.title + " Subscriptions"}
				isBlocking={true}
				closeButtonAriaLabel='Close' >
				<div>

					{errorElement }

					<DetailsList
						items={this.props.subscriptionList.subscriptions}
						setKey='set'
						checkboxVisibility={CheckboxVisibility.hidden}
						columns={this.state.columns}
						selectionMode={SelectionMode.none}
						selectionPreservedOnEmptyClick={true}
						onRenderItemColumn={this._renderItemColumn}
					/>

				</div>
				<DialogFooter>
					<Button buttonType={ButtonType.primary} onClick={this._closeNewDialog}>Ok</Button>
				</DialogFooter>
			</Dialog>
		);
	}

	private _onCancelDelete = ()  => {

			this.setState({
					selectedSubscription:this.state.selectedSubscription,
					showError:this.state.showError,
					showDeleteConfirm:false,
					deleteTarget: this.state.deleteTarget,
					subscriptionList: this.state.subscriptionList
				});	
  	}

   private _onConfirmDeleteSubscription = () => {

  		this.props.subscriptionService.deleteSubscription(this.props.subscriptionList.uniqueId, this.state.selectedSubscription.id).then( () => {

			  	this.props.removeSubscription(this.props.subscriptionList, this.state.selectedSubscription);
				var subscriptoinIndex = this.state.subscriptionList.subscriptions.indexOf(this.state.selectedSubscription);
				this.state.subscriptionList.subscriptions.splice(subscriptoinIndex,1);

				this.setState({
					subscriptionList: this.props.subscriptionList,
					selectedSubscription:this.state.selectedSubscription,
					showError:this.state.showError,
					showDeleteConfirm:false,
					deleteTarget: this.state.deleteTarget
				});	

		}).catch((error:any) =>{
				this.state.error = error;
				this.state.showError = true;
		});

}

	private _buildColumns = (): Array<IColumn> => {

		var columns: Array<IColumn> = new Array<IColumn>();

		let subscriptionUrlColumn: IColumn = {
			key: 'notificationUrl',
			name: 'Notification Url',
			fieldName: 'notificationUrl',
			minWidth: 400
		}

		columns.push(subscriptionUrlColumn);

		let expirationDateColumn: IColumn = {
			key: 'expirationDateTime',
			name: 'Expiration',
			fieldName: 'expirationDateTime',
			minWidth: 75
		}

		columns.push(expirationDateColumn);

		let clientStateColumn: IColumn = {
			key: 'clientState',
			name: 'ClientState',
			fieldName: 'clientState',
			minWidth: 100
		}

		columns.push(clientStateColumn);

		let buttonColumn: IColumn = {
			key: 'deleteButton',
			name: 'Delete',
			fieldName: 'clientState',
			minWidth: 100
		}

		columns.push(buttonColumn);

		return columns;

	}

	private _renderItemColumn = (item, index, column) => {

		let fieldContent = item[column.fieldName];
		let subscriptionItem = item as ISubscription;

		switch (column.key) {

			case 'expirationDateTime':
				var date = new Date(fieldContent);
				return <span>{date.toLocaleDateString()}</span>;

			case 'deleteButton':

				var element;

				return 	<IconButton
								 	className={styles.newLink}
									onClick={(event) => this._deleteSubscriptionRequest(subscriptionItem,event.target as HTMLElement)}
									text="Delete"
          							icon='Delete'
          							title='Delete Subscription'
          							ariaLabel='Delete Subscription' />

			default:
				return fieldContent;
		}

	}

	private _deleteSubscriptionRequest = (item:ISubscription,element:HTMLElement) => {

				this.setState({
					selectedSubscription:item,
					showError:this.state.showError,
					showDeleteConfirm:true,
					deleteTarget: element,
					subscriptionList: this.state.subscriptionList
				});	

	}

	private _closeNewDialog = () => {
		this.props.hideDialog();
	}

}
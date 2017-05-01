import * as React from 'react';
import {Dialog,DialogType,DialogFooter, Button, ButtonType } from "office-ui-fabric-react";

export interface IMessageDialogProps{
	title:string;
	message:string;
	showDialog:boolean;
	hideDialog: () => void; 
}

export interface IMessageDialogState{

}

export class MessageDialog extends React.Component< IMessageDialogProps, IMessageDialogState>{

		public render(): JSX.Element { 

			 return (
					 <Dialog
         				isOpen={ this.props.showDialog }
          				type={ DialogType.close }
          				onDismiss={ this._closeNewDialog}
          				title={this.props.title}
          				isBlocking={ true }
          				closeButtonAriaLabel='Close' >
						  		<div>
								  	<p>
										  {this.props.message}
									</p>
							 	 </div>
							<DialogFooter>
								<Button buttonType={ ButtonType.primary } onClick={this._closeNewDialog }>Ok</Button>
							</DialogFooter>
						</Dialog>
			 );
		}

		private _closeNewDialog = () => {
			this.props.hideDialog();
		}
 

}
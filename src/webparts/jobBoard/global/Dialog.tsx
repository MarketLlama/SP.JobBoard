import * as React from 'react';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';

export interface DialogPopupProps {
  primaryText : string;
  secondaryText : string;
}

export interface DialogPopupState {
  hideDialog: boolean;

}

class DialogPopup extends React.Component<DialogPopupProps, DialogPopupState> {
  constructor(props: DialogPopupProps) {
    super(props);
    this.state = {
      hideDialog: true,
    };
  }



  public setOKevent = (event : Function) =>{
    this._okEvent = event;
  }

  public render() {
    return (
      <div>
        <Dialog
          hidden={this.state.hideDialog}
          onDismiss={this._closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: this.props.primaryText,
            subText: this.props.secondaryText
          }}
          modalProps={{
            isBlocking: true,
            containerClassName: 'ms-dialogMainOverride'
          }}
        >
          <DialogFooter>
            <PrimaryButton onClick={this._okEvent} text="Ok" />
            <DefaultButton onClick={this._closeDialog} text="Cancel" />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }

  private _okEvent = null;

  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  }
}

export default DialogPopup;


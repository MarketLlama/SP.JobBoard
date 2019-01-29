import * as React from 'react';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';

export interface IErrorMessageProps {
  debug: string;
  message: string;
}

export interface IErrorMessageState {
  hidden: boolean;
}

export default class ErrorMessage extends React.Component<IErrorMessageProps, IErrorMessageState> {
  /**
   *
   */
  constructor(props) {
    super(props);
    this.state = {
      hidden: false
    };
  }

  public render() {
    let debug = null;
    if (this.props.debug) {
      debug = <pre className="alert-pre border bg-light p-2"><code>{this.props.debug}</code></pre>;
    }
    return (
      <div hidden={this.state.hidden}>
        <MessageBar messageBarType={MessageBarType.error} isMultiline={true} onDismiss={this._close} dismissButtonAriaLabel="Close" >
          <b>{this.props.message}</b>
          {debug}
        </MessageBar>
      </div>
    );
  }

  private _close = () => {
    this.setState({
      hidden: true
    });
  }
}

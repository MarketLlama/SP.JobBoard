import * as React from 'react';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';

export interface ErrorMessageProps{
  debug : string;
  message : string;
}

export default class ErrorMessage extends React.Component<ErrorMessageProps, {}> {
  /**
   *
   */
  constructor(props) {
    super(props);    
  }

  render() {
    let debug = null;
    if (this.props.debug) {
      debug = <pre className="alert-pre border bg-light p-2"><code>{this.props.debug}</code></pre>;
    }
    return (
      <MessageBar messageBarType={MessageBarType.error} isMultiline={true} dismissButtonAriaLabel="Close">
        <p className="mb-3">{this.props.message}</p>
        {debug}
      </MessageBar>
    );
  }
}
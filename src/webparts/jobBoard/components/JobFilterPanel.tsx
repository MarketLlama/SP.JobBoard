import * as React from 'react';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { ChoiceGroup } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import JobBoard from './JobBoard';

export interface JobFilterPanelProps {
  showPanel : boolean;
  parent : JobBoard;
}

export interface JobFilterPanelState {
  showPanel : boolean;
}

class JobFilterPanel extends React.Component<JobFilterPanelProps, JobFilterPanelState> {
  constructor(props: JobFilterPanelProps) {
    super(props);
    this.state = {
      showPanel: false
    };
  }

  public render(): JSX.Element {
    return (
      <div>
        <Panel
          isOpen={this.state.showPanel}
          type={PanelType.smallFixedFar}
          onDismiss={this._onClosePanel}
          headerText="Filter Panel Mockup"
          closeButtonAriaLabel="Close"
          onRenderFooterContent={this._onRenderFooterContent}
        >
          <ChoiceGroup
            options={[
              {
                key: 'A',
                text: 'Option A'
              },
              {
                key: 'B',
                text: 'Option B',
                checked: true
              },
              {
                key: 'C',
                text: 'Option C',
                disabled: true
              },
              {
                key: 'D',
                text: 'Option D',
                checked: true,
                disabled: true
              }
            ]}
            label="Pick one"
            required={true}
          />
        </Panel>
      </div>
    );
  }

  private _onClosePanel = (): void => {
    this.setState({ showPanel: false });
    this.props.parent.setState({
      showFilter : false
    });
  }

  //WARNING! To be deprecated in React v17. Use new lifecycle static getDerivedStateFromProps instead.
  public componentWillReceiveProps(nextProps: JobFilterPanelProps) {
    this.setState({
      showPanel : nextProps.showPanel
    });
  }

  private _onRenderFooterContent = (): JSX.Element => {
    return (
      <div>
        <PrimaryButton onClick={this._onClosePanel} style={{ marginRight: '8px' }}>
          Save
        </PrimaryButton>
        <DefaultButton onClick={this._onClosePanel}>Cancel</DefaultButton>
      </div>
    );
  }

  private _onShowPanel = (): void => {
    this.setState({ showPanel: true });
  }
}

export default JobFilterPanel;


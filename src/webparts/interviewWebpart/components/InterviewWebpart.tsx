import * as React from 'react';
import styles from './InterviewWebpart.module.scss';
import { IInterviewWebpartProps } from './IInterviewWebpartProps';
import { IInterviewWebpartState } from './IInterviewWebpartState';
import { Web } from "@pnp/sp";


import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';


import SubSiteCreateFormPanel from './SubSiteCreateFormPanel';
import MemberListModal from './MemberListModal';
import SubSiteList from './SubSiteList';
import SPGroup from '../models/SPGroup';

export default class InterviewWebpart extends React.Component<IInterviewWebpartProps, IInterviewWebpartState> {

  constructor(props: IInterviewWebpartProps) {
    super(props);
    // set initial state
    this.state = {
      showPanel: false,
      showModal: false,
      reloadItems: false,
    };

  }
  public componentDidMount(): void {

  }

  private _closeModal = (): void => {
    this.setState({ showModal: false });
  }

  private _onClosePanel = (): void => {
    this.setState({ showPanel: false, reloadItems: true });
  }


  private _onShowPanel = (): void => {
    this.setState({ showPanel: true });
  }

  private onShowMembersInGroups(spGroup: SPGroup, fieldName: string) {
    this.setState({ selectedGroup: spGroup, selectedFieldName: fieldName, showModal: true });
  }

  public render(): React.ReactElement<IInterviewWebpartProps> {
    return (
      <React.Fragment>
        <CommandBar
          items={[
            {
              key: 'newItem',
              name: 'Create Sub Site',
              cacheKey: 'myCacheKey', // changing this key will invalidate this items cache
              iconProps: {
                iconName: 'Add'
              },
              onClick: () => this._onShowPanel(),
              ariaLabel: 'New. Use left and right arrow keys to navigate'
            }
          ]}
        />
        <SubSiteList pageContext={this.props.pageContext} reloadItems={this.state.reloadItems} onShowMembersInGroups={this.onShowMembersInGroups.bind(this)} ></SubSiteList>
        <SubSiteCreateFormPanel pageContext={this.props.pageContext} isOpen={this.state.showPanel} onClose={this._onClosePanel.bind(this)}></SubSiteCreateFormPanel>
        <MemberListModal pageContext={this.props.pageContext} showModal={this.state.showModal} onClosed={this._closeModal.bind(this)} fieldName={this.state.selectedFieldName} spGroup={this.state.selectedGroup} ></MemberListModal>
      </React.Fragment>
    );
  }
}
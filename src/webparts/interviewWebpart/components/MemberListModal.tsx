import * as React from 'react';
import { Modal } from 'office-ui-fabric-react/lib/Modal';
import {
    DocumentCard,
    DocumentCardActivity,
    DocumentCardPreview,
    DocumentCardTitle,
    IDocumentCardPreviewProps,
    DocumentCardType
} from 'office-ui-fabric-react/lib/DocumentCard';
import User from '../models/User';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { PageContext } from "@microsoft/sp-page-context";

import SharePointService from '../services/SharePointService';
import SPGroup from '../models/SPGroup';

export interface IMemberListModalProps {
    showModal: boolean;
    spGroup?: SPGroup;
    fieldName?: string;
    onClosed: Function;
    pageContext: PageContext;
}

export interface IMemberListModalState {
    members: User[];
}

export default class MemberListModal extends React.Component<IMemberListModalProps, IMemberListModalState> {

    constructor(props: IMemberListModalProps) {
        super(props);
        this.state = {
            members: new Array<User>()
        };
    }
    public componentDidMount(): void {
        if (this.props.fieldName && this.props.spGroup) {
            this.GetMembersByGroupName(this.props.spGroup, this.props.fieldName);
        }
    }

    private GetMembersByGroupName(spGroup: SPGroup, fieldName: string) {
        console.log(spGroup[fieldName]);
        SharePointService.GetMembersByGroupName(spGroup.Url, spGroup[fieldName]).then((members) => {
            this.setState({ members });
        });
    }
    public componentWillReceiveProps(nextProps) {
        if (nextProps.fieldName && nextProps.spGroup && nextProps.showModal == true) {
            this.GetMembersByGroupName(nextProps.spGroup, nextProps.fieldName);
        }
    }
    private _closeModal = (): void => {
        this.props.onClosed();
    }

    public render() {
        const getMembers = () => {
            let renderHtml = [<div>There is no available member.</div>];
            if (this.state.members.length > 0) {
                renderHtml = this.state.members.map((member) => {
                    return <DocumentCardActivity
                        activity=""
                        people={[{ name: member.FullName, profileImageSrc: member.Image }]}
                    />;
                });
            }

            return renderHtml;

        };

        return <Modal
            titleAriaId="titleId"
            subtitleAriaId="subtitleId"
            isOpen={this.props.showModal}
            onDismiss={this._closeModal}
            isBlocking={false}
            containerClassName="ms-modalExample-container"
        >
            <div id="subtitleId" className="ms-modalExample-body">
                <DocumentCard type={DocumentCardType.compact}>
                    <div className="ms-DocumentCard-details">
                        {getMembers()}
                    </div>
                </DocumentCard>
                <DefaultButton onClick={this._closeModal.bind(this)} text="Close" />
            </div>
        </Modal>;
    }
}

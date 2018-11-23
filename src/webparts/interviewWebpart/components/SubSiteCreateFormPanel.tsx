import * as React from 'react';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { ChoiceGroup } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { PageContext } from "@microsoft/sp-page-context";

import SharePointService from '../services/SharePointService';

export interface ISubSiteCreateFormPanelProps {
    isOpen: boolean;
    onClose: Function;
    pageContext: PageContext;
}

export interface ISubSiteCreateFormPanelState {
    isLoading: boolean;
    error: {
        title: string,
        url: string,
    };
}

export default class SubSiteCreateFormPanel extends React.Component<ISubSiteCreateFormPanelProps, ISubSiteCreateFormPanelState> {

    constructor(props: ISubSiteCreateFormPanelProps) {
        super(props);

        this.state = {
            isLoading: false,
            error: {
                title: '',
                url: '',
            }
        };
    }
    public componentDidMount(): void {

    }


    private async _createSite(): Promise<void> {

        this.setState((state) => {
            state.error.title = "";
            state.error.url = "";
            return state;
        });

        const SubSiteTitle: any = this.refs.SubSiteTitle;
        const SubSiteUrlName: any = this.refs.SubSiteUrlName;
        const SubSiteDescription: any = this.refs.SubSiteDescription;
        const SubSiteLanguage: any = this.refs.SubSiteLanguage;
        const SubSiteTemplate: any = this.refs.SubSiteTemplate;

        if (SubSiteTitle.state.value === "") {
            this.setState((state) => {
                state.error.title = "Reqiured";
                return state;
            });
        }

        if (SubSiteUrlName.state.value === "") {
            this.setState((state) => {
                state.error.url = "Reqiured";
                return state;
            });
        }

        if (SubSiteTitle.state.value === "" || SubSiteUrlName.state.value === "") {
            return;
        }


        SharePointService.CreateSubSite(this.props.pageContext.web.absoluteUrl, SubSiteTitle.state.value
            , SubSiteUrlName.state.value, SubSiteDescription.state.value
            , SubSiteTemplate.state.keyChecked, SubSiteLanguage.state.keyChecked, false).then(() => {
                this.onHideLoading();
                this._onClosePanel();
                // this._getSites();
            });
        this.onShowLoading();

    }


    private _onClosePanel = (): void => {
        this.props.onClose();
    }

    private onShowLoading = (): void => {
        this.setState({ isLoading: true });

    }

    private onHideLoading = (): void => {
        this.setState({ isLoading: false });

    }

    private _onRenderFooterContent = (): JSX.Element => {
        return (
            <div>
                {
                    this.state.isLoading == true
                        ? <Spinner size={SpinnerSize.large} label="Working on it..." ariaLive="assertive" />
                        : <div>
                            <PrimaryButton onClick={this._createSite.bind(this)} style={{ marginRight: '8px' }}>
                                Save
                            </PrimaryButton>
                            <DefaultButton onClick={this._onClosePanel}>Cancel</DefaultButton>
                        </div>
                }
            </div>
        );
    }

    public render() {
        const { error } = this.state;
        return <Panel
            isOpen={this.props.isOpen}
            type={PanelType.smallFixedFar}
            onDismiss={this._onClosePanel}
            headerText="Create Sub Site"
            closeButtonAriaLabel="Close"
            onRenderFooterContent={this._onRenderFooterContent}
        >

            <TextField label="Title" required={true} ref="SubSiteTitle" errorMessage={error.title} />
            <TextField label="Description" multiline rows={4} ref="SubSiteDescription" />
            <TextField label="URL Name" required={true} ref="SubSiteUrlName" errorMessage={error.url} />
            <ChoiceGroup
                ref="SubSiteLanguage"
                label="Select a language:"
                placeholder="Select a language"
                defaultSelectedKey="1033"
                options={[
                    { key: '1033', text: 'English', checked: true },
                    { key: '1045', text: 'Polish' },
                ]}
            />
            <ChoiceGroup
                ref="SubSiteTemplate"
                options={[
                    {
                        key: 'WIKI#0',
                        text: 'Wiki Site'
                    },
                    {
                        key: 'STS#3',
                        text: 'Team Site',
                        checked: true
                    },
                    {
                        key: 'BLOG#0',
                        text: 'Blog',
                    },
                    {
                        key: 'BLANKINTERNETCONTAINER#0',
                        text: 'Publishing Portal',
                    }
                ]}
                label="Select a template:"
                required={true}
            />
        </Panel>;
    }
}

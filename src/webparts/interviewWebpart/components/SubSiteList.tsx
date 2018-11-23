import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, Selection, IColumn, IDetailsList, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { PageContext } from "@microsoft/sp-page-context";

import { TextField } from 'office-ui-fabric-react/lib/TextField';
import SPGroup from '../models/SPGroup';
import SharePointService from '../services/SharePointService';

export interface ISubSiteListProps {
    reloadItems?: boolean;
    onShowMembersInGroups: Function;
    pageContext: PageContext;
}

export interface ISubSiteListState {
    items: SPGroup[];
    allItems: SPGroup[];
    columns: IColumn[];
}


export default class SubSiteList extends React.Component<ISubSiteListProps, ISubSiteListState> {
    private _selection: Selection;

    constructor(props: ISubSiteListProps) {
        super(props);
        this.state = {
            items: new Array<SPGroup>(),
            allItems: new Array<SPGroup>(),
            columns: [
                {
                    key: 'Title',
                    name: 'Title',
                    fieldName: 'Title',
                    minWidth: 100,
                    maxWidth: 200,
                    isRowHeader: true,
                    isResizable: true,
                    isSorted: true,
                    isSortedDescending: false,
                    sortAscendingAriaLabel: 'Sorted A to Z',
                    sortDescendingAriaLabel: 'Sorted Z to A',
                    onColumnClick: this._onColumnClick,
                    data: 'string',
                    isPadded: true
                },
                {
                    key: 'Owners',
                    name: 'Owners',
                    fieldName: 'Owners',
                    minWidth: 100,
                    maxWidth: 200,
                    isResizable: true,
                    isSorted: true,
                    sortAscendingAriaLabel: 'Sorted A to Z',
                    sortDescendingAriaLabel: 'Sorted Z to A',
                    data: 'string',
                    ariaLabel: 'Operations for Owners',
                    onColumnClick: this._onColumnClick,

                },
                {
                    key: 'Members',
                    name: 'Members',
                    fieldName: 'Members',
                    minWidth: 100,
                    maxWidth: 200,
                    isResizable: true,
                    isSorted: true,
                    sortAscendingAriaLabel: 'Sorted A to Z',
                    sortDescendingAriaLabel: 'Sorted Z to A',
                    ariaLabel: 'Operations for Members',
                    data: 'string',
                    onColumnClick: this._onColumnClick,

                },
                {
                    key: 'Visitors',
                    name: 'Visitors',
                    fieldName: 'Visitors',
                    minWidth: 100,
                    maxWidth: 200,
                    isResizable: true,
                    isSorted: true,
                    sortAscendingAriaLabel: 'Sorted A to Z',
                    sortDescendingAriaLabel: 'Sorted Z to A',
                    ariaLabel: 'Operations for Visitors',
                    data: 'string',
                    onColumnClick: this._onColumnClick,

                }
            ]
        };

        this._getSites.bind(this);
    }
    public componentDidMount(): void {
        this._getSites();
    }

    public componentWillReceiveProps(nextProps) {
        if (nextProps.reloadItems == true) {
            this._getSites();
        }
    }
    private async _getSites() {
        SharePointService.GetAllSites(this.props.pageContext.web.absoluteUrl).then((response) => {
            const items: any = response;
            // set our Component´s State 
            this.setState({ items, allItems: items });
        }).catch((errors) => {

        });


    }

    public handleChange = (event) => {
        const { target: { name, value } } = event;
    }
    private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
        const { columns, items } = this.state;
        let newItems: SPGroup[] = items.slice();
        const newColumns: IColumn[] = columns.slice();
        const currColumn: IColumn = newColumns.filter((currCol: IColumn, idx: number) => {
            return column.key === currCol.key;
        })[0];
        newColumns.forEach((newCol: IColumn) => {
            if (newCol === currColumn) {
                currColumn.isSortedDescending = !currColumn.isSortedDescending;
                currColumn.isSorted = true;
            } else {
                newCol.isSorted = false;
                newCol.isSortedDescending = true;
            }
        });
        newItems = this._sortItems(newItems, currColumn.fieldName || '', currColumn.isSortedDescending);
        this.setState({
            columns: newColumns,
            items: newItems
        });
    }
    private _sortItems = (items: SPGroup[], sortBy: string, descending = false): SPGroup[] => {
        if (descending) {
            return items.sort((a: SPGroup, b: SPGroup) => {
                if (a[sortBy] < b[sortBy]) {
                    return 1;
                }
                if (a[sortBy] > b[sortBy]) {
                    return -1;
                }
                return 0;
            });
        } else {
            return items.sort((a: SPGroup, b: SPGroup) => {
                if (a[sortBy] < b[sortBy]) {
                    return -1;
                }
                if (a[sortBy] > b[sortBy]) {
                    return 1;
                }
                return 0;
            });
        }
    }
    private _onChangeText = (text): void => {
        const items = this.state.items;
        const allItems = this.state.allItems;
        this.setState({ items: text ? allItems.filter(i => i.Title.toLowerCase().indexOf(text) > -1) : allItems });
    }

    private onShowMembersInGroups(spGroup: SPGroup, fieldName: string) {
        this.props.onShowMembersInGroups(spGroup, fieldName);
    }

    private _onRenderItemColumn(item: any, index: number, column: IColumn): JSX.Element {
        if (column.fieldName === 'Title') {
            return <Link target="_blank" href={item.Url}>{item[column.fieldName]}</Link>;
        }
        else {
            return <Link onClick={this.onShowMembersInGroups.bind(this, item, column.fieldName)}>{item[column.fieldName]}</Link>;
        }
    }

    public render() {
        const { items, columns } = this.state;

        return <div>
            <TextField label="Filter by name:" name="filter" onChanged={this._onChangeText} />

            <MarqueeSelection selection={this._selection}>
                <DetailsList
                    onRenderItemColumn={this._onRenderItemColumn.bind(this)}
                    items={items}
                    columns={columns}
                    setKey="set"
                    selectionMode={SelectionMode.none}
                    layoutMode={DetailsListLayoutMode.fixedColumns}
                    selection={this._selection}
                    selectionPreservedOnEmptyClick={true}
                    ariaLabelForSelectionColumn="Toggle selection"
                    ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                />
            </MarqueeSelection>
        </div>;
    }
}

import * as React from 'react';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { getId } from 'office-ui-fabric-react/lib/Utilities';
import { hiddenContentStyle, mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { ContextualMenu } from 'office-ui-fabric-react/lib/ContextualMenu';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import * as $ from 'jquery';
import pnp, { List, App, Web, CamlQuery, EmailProperties } from "sp-pnp-js";
import { IAssignPropertyDialogProps } from './IAssignPropertyDialogProps';
import { IAssignPropertyDialogState } from './IAssignPropertyDialogState';
import { IDropdownStyles, IDropdownOption, Dropdown, SelectedPeopleList } from 'office-ui-fabric-react';
import { sp } from '@pnp/sp';

const screenReaderOnly = mergeStyles(hiddenContentStyle);

let selectDocs: any[] = [];
let selectUser: any[] = [];

let cntDoc = 0;

let docArrlength;
let userArrlength;

let requestDigest;

let dropdownStyles: Partial<IDropdownStyles> = {
    dropdown: { width: 300 }
};

let GroupOption: IDropdownOption[] = [];
let JobRoleOption: IDropdownOption[] = [];
let DepartOption: IDropdownOption[] = [];
let ProjectOption: IDropdownOption[] = [];

export class AssignProperty extends React.Component<IAssignPropertyDialogProps, IAssignPropertyDialogState> {
    // Use getId() to ensure that the IDs are unique on the page.
    // (It's also okay to use plain strings without getId() and manually ensure uniqueness.)
    private _labelId: string = getId('dialogLabel');
    private _subTextId: string = getId('subTextLabel');
    private _dragOptions = {
        moveMenuItemText: 'Move',
        closeMenuItemText: 'Close',
        menu: ContextualMenu
    };

    public constructor(props: IAssignPropertyDialogProps) {
        super(props);
        this.updateDocument = this.updateDocument.bind(this);
        this._updateDocumentPropertyDigest = this._updateDocumentPropertyDigest.bind(this);
        this.state = {
            hidePropertyDialog: true,
            selectedGroupItem: [],
            selectedGroupItemText: [],
            selectedDepItem: [],
            selectedProjectItem: [],
            selectedJobRoleItem: [],
            selectedDepItemText: [],
            selectedProjectItemText: [],
            selectedJobRoleItemText: []
        };
    }

    public async componentWillReceiveProps(nextProps: IAssignPropertyDialogProps) {
        console.log(nextProps);

        this.getItems();
        this.setState({
            selectedGroupItem: [],
            selectedDepItem: [],
            selectedProjectItem: [],
            selectedJobRoleItem: [],
            hidePropertyDialog: nextProps.hidePropertyDialog
        });
        selectDocs = nextProps.selectDocs;
    }

    public async getItems() {
        await $.ajax({
            url: "https://sjch.sharepoint.com/sites/SharedCentre/_api/web/lists/getbytitle('Country')/items",
            method: "Get",
            headers: { "Accept": "application/json; odata=verbose" },
            success: (data) => {
                GroupOption = [];
                for (let item of data.d.results) {
                    let temp = { key: item.ID, text: item.Title };
                    GroupOption.push(temp);
                }
            },
            error: (data, errorCode, errorMessage) => {
                alert(errorMessage);
            }
        });

        await $.ajax({
            url: "https://sjch.sharepoint.com/sites/SharedCentre/_api/web/lists/getbytitle('JobRoles')/items",
            method: "Get",
            headers: { "Accept": "application/json; odata=verbose" },
            success: (data) => {
                JobRoleOption = [];
                for (let item of data.d.results) {
                    let temp = { key: item.ID, text: item.Title };
                    //JobRoleOption.push(temp);
                }
            },
            error: (data, errorCode, errorMessage) => {
                alert(errorMessage);
            }
        });

        await $.ajax({
            url: "https://sjch.sharepoint.com/sites/SharedCentre/_api/web/lists/getbytitle('Department')/items",
            method: "Get",
            headers: { "Accept": "application/json; odata=verbose" },
            success: (data) => {
                DepartOption = [];
                for (let item of data.d.results) {
                    let temp = { key: item.ID, text: item.Title };
                    //DepartOption.push(temp);
                }
            },
            error: (data, errorCode, errorMessage) => {
                alert(errorMessage);
            }
        });

        await $.ajax({
            url: "https://sjch.sharepoint.com/sites/SharedCentre/_api/web/lists/getbytitle('SiteProjects')/items",
            method: "Get",
            headers: { "Accept": "application/json; odata=verbose" },
            success: (data) => {
                ProjectOption = [];
                for (let item of data.d.results) {
                    let temp = { key: item.ID, text: item.Title };
                    //ProjectOption.push(temp);
                }
            },
            error: (data, errorCode, errorMessage) => {
                alert(errorMessage);
            }
        });

    }

    private _onGroupChange = async (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): Promise<void> => {
        const newSelectedItems = [...this.state.selectedGroupItem];
        const newSelectedItemsText = [...this.state.selectedGroupItemText];
        if (item.selected) {
            // add the option if it's checked
            newSelectedItems.push(item.key as number);
            newSelectedItemsText.push(item.text as string);
        } else {
            // remove the option if it's unchecked
            const currIndex = newSelectedItems.indexOf(item.key as number);
            if (currIndex > -1) {
                newSelectedItems.splice(currIndex, 1);
            }

            // remove the option if it's unchecked
            const currIndextest = newSelectedItemsText.indexOf(item.text as string);
            if (currIndextest > -1) {
                newSelectedItemsText.splice(currIndextest, 1);
            }
        }
        this.setState({
            selectedGroupItem: newSelectedItems,
            selectedGroupItemText: newSelectedItemsText
        });

        var projectVal = "";
        $.each(newSelectedItemsText, (e, val) => {
            projectVal += "<Value Type=\"Lookup\">" + val + "</Value>";
        });
        let caml = "";
        caml = "<Where>" +
            "<In><FieldRef Name=\"Country\" /><Values>" + projectVal +
            "</Values></In>" +
            "</Where>";

        let camlQuery: string = '';
        camlQuery = `<View>
                                <Query>
                                    <OrderBy>
                                        <FieldRef Name="ID" 'Ascending="FALSE"'}/> 
                                    </OrderBy> 
                                      ${caml}
                                    </Query>
                            <ViewFields>
                                <FieldRef Name="ID"/>
                                <FieldRef Name="Title"/>
                                <FieldRef Name="Country"/>                     
                            </ViewFields>`;

        const countQuery: CamlQuery = {
            ViewXml: `${camlQuery}</View>`,
        };

        let web = new Web("https://sjch.sharepoint.com/sites/SharedCentre/");

        const result = await web.lists.getByTitle("SiteProjects").getItemsByCAMLQuery(countQuery, 'FieldValuesAsText');

        let option = [];
        option.push({
            key: "",
            text: "Select an Option",
        });
        result.forEach((item: any) => {
            option.push({
                key: item.Id,
                text: item.Title,
            });
        });
        ProjectOption = option;
        this.setState({
            selectedGroupItemText: newSelectedItemsText
        });

    }

    private _onProjectChange = async (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): Promise<void> => {
        const newSelectedItems = [...this.state.selectedProjectItem];
        const newSelectedItemsText = [...this.state.selectedProjectItemText];
        if (item.selected) {
            // add the option if it's checked
            newSelectedItems.push(item.key as number);
            newSelectedItemsText.push(item.text as string);
        } else {
            // remove the option if it's unchecked
            const currIndex = newSelectedItems.indexOf(item.key as number);
            if (currIndex > -1) {
                newSelectedItems.splice(currIndex, 1);
            }

            const currIndextest = newSelectedItemsText.indexOf(item.text as string);
            if (currIndextest > -1) {
                newSelectedItemsText.splice(currIndextest, 1);
            }
        }
        this.setState({
            selectedProjectItem: newSelectedItems,
            selectedProjectItemText: newSelectedItemsText
        });

        var Val = "";
        $.each(newSelectedItemsText, (e, val) => {
            Val += "<Value Type=\"Lookup\">" + val + "</Value>";
        });
        let caml = "";
        caml = "<Where>" +
            "<In><FieldRef Name=\"Site_x0020_Projects\" /><Values>" + Val +
            "</Values></In>" +
            "</Where>";

        let camlQuery: string = '';
        camlQuery = `<View>
                                <Query>
                                    <OrderBy>
                                        <FieldRef Name="ID" 'Ascending="FALSE"'}/> 
                                    </OrderBy> 
                                      ${caml}
                                    </Query>
                            <ViewFields>
                                <FieldRef Name="ID"/>
                                <FieldRef Name="Title"/>
                                <FieldRef Name="Site_x0020_Projects"/>                     
                            </ViewFields>`;

        const countQuery: CamlQuery = {
            ViewXml: `${camlQuery}</View>`,
        };

        let web = new Web("https://sjch.sharepoint.com/sites/SharedCentre/");

        const result = await web.lists.getByTitle("Department").getItemsByCAMLQuery(countQuery, 'FieldValuesAsText');

        let option = [];
        option.push({
            key: "",
            text: "Select an Option",
        });
        result.forEach((item: any) => {
            option.push({
                key: item.Id,
                text: item.Title,
            });
        });
        DepartOption = option;
        this.setState({
            selectedProjectItemText: newSelectedItemsText
        });

    }

    private _onDepChange = async (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): Promise<void> => {
        const newSelectedItems = [...this.state.selectedDepItem];
        const newSelectedItemsText = [...this.state.selectedDepItemText];
        if (item.selected) {
            // add the option if it's checked
            newSelectedItems.push(item.key as number);
            newSelectedItemsText.push(item.text as string);
        } else {
            // remove the option if it's unchecked 
            const currIndex = newSelectedItems.indexOf(item.key as number);
            if (currIndex > -1) {
                newSelectedItems.splice(currIndex, 1);
            }

            const currIndextest = newSelectedItemsText.indexOf(item.text as string);
            if (currIndextest > -1) {
                newSelectedItemsText.splice(currIndextest, 1);
            }
        }
        this.setState({
            selectedDepItem: newSelectedItems,
            selectedDepItemText: newSelectedItemsText
        });

        var Val = "";
        $.each(newSelectedItemsText, (e, val) => {
            Val += "<Value Type=\"Lookup\">" + val + "</Value>";
        });
        let caml = "";
        caml = "<Where>" +
            "<In><FieldRef Name=\"DEPARTMENT\" /><Values>" + Val +
            "</Values></In>" +
            "</Where>";

        let camlQuery: string = '';
        camlQuery = `<View>
                                <Query>
                                    <OrderBy>
                                        <FieldRef Name="ID" 'Ascending="FALSE"'}/> 
                                    </OrderBy> 
                                      ${caml}
                                    </Query>
                            <ViewFields>
                                <FieldRef Name="ID"/>
                                <FieldRef Name="Title"/>
                                <FieldRef Name="DEPARTMENT"/>                     
                            </ViewFields>`;

        const countQuery: CamlQuery = {
            ViewXml: `${camlQuery}</View>`,
        };

        let web = new Web("https://sjch.sharepoint.com/sites/SharedCentre/");

        const result = await web.lists.getByTitle("JobRoles").getItemsByCAMLQuery(countQuery, 'FieldValuesAsText');

        let option = [];
        option.push({
            key: "",
            text: "Select an Option",
        });
        result.forEach((item: any) => {
            option.push({
                key: item.Id,
                text: item.Title,
            });
        });
        JobRoleOption = option;
        this.setState({
            selectedDepItemText: newSelectedItemsText
        });
    }

    private _onRoleChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
        const newSelectedItems = [...this.state.selectedJobRoleItem];
        const newSelectedItemsText = [...this.state.selectedJobRoleItemText];
        if (item.selected) {
            // add the option if it's checked
            newSelectedItems.push(item.key as number);
            newSelectedItemsText.push(item.text as string);
        } else {
            // remove the option if it's unchecked 
            const currIndex = newSelectedItems.indexOf(item.key as number);
            if (currIndex > -1) {
                newSelectedItems.splice(currIndex, 1);
            }

            const currIndextest = newSelectedItemsText.indexOf(item.text as string);
            if (currIndextest > -1) {
                newSelectedItemsText.splice(currIndextest, 1);
            }
        }
        this.setState({
            selectedJobRoleItem: newSelectedItems,
            selectedJobRoleItemText: newSelectedItemsText
        });
    }

    private _closeDialog = (): void => {
        this.setState({ hidePropertyDialog: true });
    }

    public render() {
        const { selectedGroupItem } = this.state;
        const { selectedJobRoleItem } = this.state;
        const { selectedDepItem } = this.state;
        const { selectedProjectItem } = this.state;
        return (
            <div>
                <label id={this._labelId} className={screenReaderOnly}>
                    My sample Label
    </label>
                <label id={this._subTextId} className={screenReaderOnly}>
                    My Sample description
    </label>

                <Dialog
                    hidden={this.state.hidePropertyDialog}
                    onDismiss={this._closeDialog}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Update Document Property',
                        closeButtonAriaLabel: 'Close',
                        subText: 'Please select Group and Job Roles to Share selected Documents.'
                    }}
                    modalProps={{
                        titleAriaId: this._labelId,
                        subtitleAriaId: this._subTextId,
                        isBlocking: false,
                        styles: { main: { maxWidth: 450 } },
                        dragOptions: undefined
                    }}
                >
                    <Dropdown
                        placeholder="Select options"
                        label="Select Country"
                        selectedKeys={selectedGroupItem}
                        onChange={this._onGroupChange}
                        multiSelect
                        options={GroupOption}
                        styles={dropdownStyles}
                    />

                    <Dropdown
                        placeholder="Select options"
                        label="Select Site Projects"
                        selectedKeys={selectedProjectItem}
                        onChange={this._onProjectChange}
                        multiSelect
                        options={ProjectOption}
                        styles={dropdownStyles}
                    />

                    <Dropdown
                        placeholder="Select options"
                        label="Select Department"
                        selectedKeys={selectedDepItem}
                        onChange={this._onDepChange}
                        multiSelect
                        options={DepartOption}
                        styles={dropdownStyles}
                    />

                    <Dropdown
                        placeholder="Select options"
                        label="Select Job Roles"
                        selectedKeys={selectedJobRoleItem}
                        onChange={this._onRoleChange}
                        multiSelect
                        options={JobRoleOption}
                        styles={dropdownStyles}
                    />

                    <DialogFooter>
                        <PrimaryButton onClick={this._updateDocumentPropertyDigest} text="Update" />
                        <DefaultButton onClick={this._closeDialog} text="Cancel" />
                    </DialogFooter>

                </Dialog>
            </div>
        );
    }

    private _updateDocumentPropertyDigest = async (): Promise<void> => {
        this.setState({ hidePropertyDialog: true });

        console.log(this.state.selectedGroupItem);
        console.log(this.state.selectedJobRoleItem);
        console.log(this.props.selectDocs);
        console.log(this.state.selectedDepItem);
        console.log(this.state.selectedProjectItem);

        cntDoc = 0;
        docArrlength = selectDocs.length;

        await $.ajax({
            url: "https://sjch.sharepoint.com/sites/SharedCentre/_api/contextinfo",
            method: "POST",
            headers: { "Accept": "application/json; odata=verbose" },
            success: (data) => {
                requestDigest = data.d.GetContextWebInformation.FormDigestValue;
                this.updateDocument(selectDocs[cntDoc], requestDigest);
            },
            error: (data, errorCode, errorMessage) => {
                alert(errorMessage);
            }
        });
    }

    private updateDocument = async (doc: any, requestDigest: any): Promise<void> => {

        console.log(doc);

        var itemProperties = {
            "__metadata": { "type": "SP.Data.Shared_x0020_DocumentsItem" }
        };

        if (this.state.selectedGroupItem.length > 0) {
            itemProperties["GroupsId"] = {
                '__metadata': { type: 'Collection(Edm.Int32)' },
                'results': this.state.selectedGroupItem
            };
        }
        if (this.state.selectedProjectItem.length > 0) {
            itemProperties["SiteProjectsId"] = {
                '__metadata': { type: 'Collection(Edm.Int32)' },
                'results': this.state.selectedProjectItem
            };
        }

        if (this.state.selectedDepItem.length > 0) {
            itemProperties["DepartmentId"] = {
                '__metadata': { type: 'Collection(Edm.Int32)' },
                'results': this.state.selectedDepItem
            };
        }

        if (this.state.selectedJobRoleItem.length > 0) {
            itemProperties["JobRoleId"] = {
                '__metadata': { type: 'Collection(Edm.Int32)' },
                'results': this.state.selectedJobRoleItem
            };
        }

        await $.ajax({
            url: "https://sjch.sharepoint.com/sites/SharedCentre/_api/web/lists/getbytitle('Documents')/items(" + doc["key"] + ")",
            method: "POST",
            contentType: "application/json;odata=verbose",
            data: JSON.stringify(itemProperties),
            headers: {
                "Accept": "application/json;odata=verbose",
                "X-RequestDigest": requestDigest,
                "X-HTTP-Method": "MERGE",
                "If-Match": '*'
            },
            success: (data) => {
                console.log(data);
                cntDoc++;
                if (cntDoc < docArrlength) {
                    this.updateDocument(selectDocs[cntDoc], requestDigest);
                }
                else {
                    this.props.callback();
                }
            },
            error: (data) => {
                console.log(data);
            }
        });
    }

    private _showDialog = (): void => {
        this.setState({ hidePropertyDialog: false });
    }
}
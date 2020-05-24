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
import { IAssignDialogProps } from './IAssignDialogProps';
import { IAssignDialogState } from './IAssignDialogState';

const screenReaderOnly = mergeStyles(hiddenContentStyle);

let selectDocs: any[] = [];
let selectUser: any[] = [];

let cntDoc = 0;
let cntUser = 0;

let docArrlength;
let userArrlength;

let requestDigest;
export class AssignDialog extends React.Component<IAssignDialogProps, IAssignDialogState> {
    // Use getId() to ensure that the IDs are unique on the page.
    // (It's also okay to use plain strings without getId() and manually ensure uniqueness.)
    private _labelId: string = getId('dialogLabel');
    private _subTextId: string = getId('subTextLabel');
    private _dragOptions = {
        moveMenuItemText: 'Move',
        closeMenuItemText: 'Close',
        menu: ContextualMenu
    };

    public constructor(props: IAssignDialogProps) {
        super(props);

        this.assignDocumentUser = this.assignDocumentUser.bind(this);

        this.state = {
            hideDialog: true,
            
        };
    }

    public async componentWillReceiveProps(nextProps: IAssignDialogProps) {
        console.log(nextProps);
        this.setState({
            hideDialog: nextProps.hideDialog
        });
        selectDocs = nextProps.selectDocs;
    }

    public render() {

        return (
            <div>
                <label id={this._labelId} className={screenReaderOnly}>
                    My sample Label
        </label>
                <label id={this._subTextId} className={screenReaderOnly}>
                    My Sample description
        </label>

                <Dialog
                    hidden={this.state.hideDialog}
                    onDismiss={this._closeDialog}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Give Document Permission',
                        closeButtonAriaLabel: 'Close',
                        subText: 'Please select People/Group to Share selected Documents.'
                    }}
                    modalProps={{
                        titleAriaId: this._labelId,
                        subtitleAriaId: this._subTextId,
                        isBlocking: false,
                        styles: { main: { maxWidth: 450 } },
                        dragOptions: undefined
                    }}
                >
                    <PeoplePicker
                        context={this.props.context}
                        titleText=""
                        personSelectionLimit={100}
                        showtooltip={true}
                        isRequired={false}
                        disabled={false}
                        ensureUser={true}
                        showHiddenInUI={false}
                        principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup]}
                        selectedItems={this._getContentPeoplePickerItem}
                        defaultSelectedUsers={undefined}
                        resolveDelay={1000}
                    />
                    <DialogFooter>
                        <PrimaryButton onClick={this._shareDocument} text="Share" />
                        <DefaultButton onClick={this._closeDialog} text="Don't Share" />
                    </DialogFooter>

                </Dialog>
            </div>
        );
    }

    private _getContentPeoplePickerItem = (items: any): void => {
        selectUser = [];
        selectUser.push(items);
    }

    private _shareDocument = async (): Promise<void> => {

        if (selectUser.length > 0) {
            cntDoc = 0;
            cntUser = 0;
            docArrlength = selectDocs.length;
            userArrlength = selectUser[0].length;


            await $.ajax({
                url: "https://sjch.sharepoint.com/sites/SharedCentre/_api/contextinfo",
                method: "POST",
                headers: { "Accept": "application/json; odata=verbose" },
                success: (data) => {
                    requestDigest = data.d.GetContextWebInformation.FormDigestValue;
                    this.assignDocumentUser(selectDocs[cntDoc], selectUser[0][cntUser]);
                },
                error: (data, errorCode, errorMessage) => {
                    alert(errorMessage);
                }
            });
        }
        else {
            alert("Please select User or group");
        }

    }


    private assignDocumentUser = async (doc: any, user: any): Promise<void> => {
        console.log(doc.FilePath);
        console.log(user.id);

        //let userDetail = await pnp.sp.web.siteUsers.getById(user.id).get();
        //let userEmail = userDetail.Email;
        let web = new Web("https://sjch.sharepoint.com/sites/SharedCentre/");
        const file = web.getFileByServerRelativePath(doc.FilePath);
        const fileItem = await file.getItem();

        await fileItem.breakRoleInheritance(true);

        let restSource = "https://sjch.sharepoint.com/sites/SharedCentre/_api/web/GetFileByServerRelativeUrl('" + doc.FilePath + "')/ListItemAllFields/roleassignments/addroleassignment(principalid=" + parseInt(user.id) + ",roledefid=1073741826)";

        $.ajax({
            url: restSource,
            method: "POST",
            headers: {
                'accept': 'application/json;odata=verbose',
                'content-type': 'application/json;odata=verbose',
                'X-RequestDigest': requestDigest
            },
            success: (data) => {

                cntUser++;

                if (cntUser < userArrlength) {
                    this.assignDocumentUser(selectDocs[cntDoc], selectUser[0][cntUser]);
                }
                else {
                    cntDoc++;
                    if (cntDoc < docArrlength) {
                        cntUser = 0;
                        this.assignDocumentUser(selectDocs[cntDoc], selectUser[0][cntUser]);
                    }
                    else {
                        this.props.callback();
                    }

                }

                // const emailProps: EmailProperties = {
                //     To: [userEmail],
                //     Subject: "New Document Assigned Please Review",
                //     Body: "Please review Shared Email",
                // };
                // pnp.sp.utility.sendEmail(emailProps).then(_ => {

                // });

            },
            error: (data, errorCode, errorMessage) => {
                alert(errorMessage);
            }
        });

    }

    public getRequestDigest(requestDigest) {

    }
    private _showDialog = (): void => {
        this.setState({ hideDialog: false });
    }

    private _closeDialog = (): void => {
        this.setState({ hideDialog: true });
    }
}

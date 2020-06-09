import * as React from 'react';
import styles from './DocumentTable.module.scss';
import { IDocumentTableProps } from './IDocumentTableProps';
import { IDocumentTableState } from './IDocumentTableState';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
/*import styles from './ReactPnpResponsiveDataTable.module.scss';*/
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as $ from 'jquery';
import 'jszip/dist/jszip';
import 'pdfmake/build/pdfmake';
import 'datatables.net';
import 'datatables.net-responsive';
import 'datatables.net-buttons';
import * as FileSaver from 'file-saver';
import 'datatables.net-buttons/js/buttons.html5';
import 'datatables.net-buttons/js/buttons.print';
import { IDocument } from './IDocumentTableState';
import { IService } from '../../../Services/IService';
import { IframeDialog } from '../../../Components/IFrameDialog/IframeDialog';
import { AssignDialog } from './../../../components/assignDialog/AssignUser';
import { AssignProperty } from './../../../components/AssignPropertyDialog/AssignProperty';
import { Image } from 'office-ui-fabric-react/lib/Image';
import pnp, { List, App, Web, CamlQuery } from "sp-pnp-js";
import ReactHtmlParser, { processNodes, convertNodeToElement, htmlparser2 } from 'react-html-parser';
//import { ISpfxPdfProps } from './ISpfxPdfProps';
import * as jsPDF from 'jspdf';
import html2canvas from 'html2canvas';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
require('./tablestyle.css');
SPComponentLoader.loadCss('https://cdn.datatables.net/responsive/2.2.3/css/responsive.bootstrap.min.css');
SPComponentLoader.loadCss('https://cdn.datatables.net/1.10.20/css/jquery.dataTables.min.css');
SPComponentLoader.loadCss('https://cdn.datatables.net/buttons/1.6.0/css/buttons.dataTables.min.css');
SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css');
SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.53/vfs_fonts.js');

SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css');

import { CSVLink } from "react-csv";

import { PrimaryButton, Dropdown, IDropdownOption, IDropdownStyles, DefaultButton, Dialog, DialogType, DialogFooter, TextField } from 'office-ui-fabric-react';
let items: any[] = [];
let selectDocs: any[] = [];
let enteredText = "";

let isEditDialogShow = true;
let isDialogShow = true;
let itemURl = "";
let country = "";
let optionsGroup: IDropdownOption[] = [];
let Grpoptions: IDropdownOption[] = [];

let optionsJob: IDropdownOption[] = [];
let TempoptionsJob: IDropdownOption[] = [];
let siteproject;
let siteprojectJob: IDropdownOption[] = [];
let TempsiteprojectJob: IDropdownOption[] = [];

let Langoptions: IDropdownOption[] = [];

let jobRoleOption: any = {};
let jobRoleOptionArr: IDropdownOption[] = [];
let tempJobRole: IDropdownOption[] = [];

let dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300, marginBottom: "10px" }
};

let GroupArr = [];
let JobArr = [];
let table;
let Dep = [];
let jobRole = [];
let isAdmin = false;
let depChecked = false;
let rebindDrop = false;
let categoryCheck: any = [];

export default class DocumentTable extends React.Component<IDocumentTableProps, IDocumentTableState> {

  private _selection: Selection;
  private _allItems: IDocument[];
  private spHttpClient: SPHttpClient;

  constructor(props: IDocumentTableProps) {
    super(props);


    this.generateDocuments = this.generateDocuments.bind(this);
    this.openProductDialog = this.openProductDialog.bind(this);
    this.assignDocument = this.assignDocument.bind(this);
    this.callback = this.callback.bind(this);
    this.assignProperties = this.assignProperties.bind(this);
    this.getjobRoles = this.getjobRoles.bind(this);
    this._onChange = this._onChange.bind(this);

    this.state = {
      items: this._allItems,
      //columns: columns,
      //selectionDetails: this._getSelectionDetails(),
      isModalSelection: false,
      isCompactMode: false,
      announcedMessage: undefined,
      hideDialog: true,
      hideModal: true,
      hidePropertyDialog: true,
      GroupItem: [],
      JobRoleItem: [],
      GroupArr: [],
      siteProjectArr: [],
      TsiteProjectArr: [],
      JobArr: [],
      LangArr: [],
      TLangArr: [],
      TGroupArr: [],
      TJobArr: [],
      TJobArrNew: [],
      hidePrintDialog: true,
      Text: "",
      depChecked: false,
      jobR: [],
      TjobR: []
    };

  }

  public async componentDidMount(): Promise<void> {
    selectDocs = [];

    this.getDropdownItems();
    await this.getjobRoles(this.props.site, this.props.currentUser);
    console.log(Dep);
    console.log(jobRole);
    if (jobRole.indexOf("Head Office CEO") > -1 || jobRole.indexOf("Head Office Business Process Specialist") > -1 || jobRole.indexOf("Head Office Document Controller") > -1) {
      isAdmin = true;
    }
    else {
      isAdmin = false;
    }

    this.generateDocuments(this.props.site, "Documents", this.props.currentUser, this.props, this.state.TGroupArr, this.state.TJobArr, this.state.TLangArr, this.state.TsiteProjectArr);
  }

  public async getjobRoles(siteUrl: string, currentUser: string) {
    let caml = "";
    let camlQuery: string = '';
    camlQuery = `<View Scope='Recursive'>
                        <Query>
                            ${caml}
							<OrderBy>
								<FieldRef Name="ID" Ascending='TRUE' />
                            </OrderBy>
                            </Query>
					<ViewFields>
                        <FieldRef Name="ID"/>
                        <FieldRef Name="Title"/>
                        <FieldRef Name="DEPARTMENT"/>
                        <FieldRef Name="Users"/>
                    </ViewFields>`;

    const countQuery: CamlQuery = {
      ViewXml: `${camlQuery}</View>`,
    };

    let web = new Web(siteUrl);

    const items: any[] = await web.lists.getByTitle("JobRoles").getItemsByCAMLQuery(countQuery, 'FieldValuesAsText');
    console.log(items);

    Dep = [];
    jobRole = [];
    $.each(items, function (e, val) {
      var users = val.FieldValuesAsText.Users;
      var cu = currentUser;
      var userArr = users.indexOf(';') > -1 ? users.split(';') : users;
      if (users.indexOf(';') > -1 == true) {
        if (users.indexOf(cu) > -1) {
          var DepArr = val.FieldValuesAsText.DEPARTMENT.trim();

          var temp = DepArr.indexOf(';') > -1 ? DepArr.split(';') : DepArr;
          var temp1 = val.FieldValuesAsText.Title;
          if (DepArr.indexOf(';') > -1) {
            $.each(temp, function (e, val) { Dep.push(val.trim()); });
          }
          else {
            Dep.push(temp);
          }
          jobRole.push(temp1);
        }

      }
      else {
        if (userArr == cu) {
          console.log(val.FieldValuesAsText);
          var DepArr = val.FieldValuesAsText.DEPARTMENT.trim();
          var temp = DepArr.indexOf(';') > -1 ? DepArr.split(';') : DepArr;
          var temp1 = val.FieldValuesAsText.Title;
          if (DepArr.indexOf(';') > -1) {
            $.each(temp, function (e, val) { Dep.push(val.trim()); });
          }
          else {
            Dep.push(temp);
          }
          jobRole.push(temp1);
        }
      }
    });

    if (jobRole.indexOf("Head Office CEO") > -1 || jobRole.indexOf("Head Office Business Process Specialist") > -1 || jobRole.indexOf("Head Office Document Controller") > -1) {
      isAdmin = true;
    }
    else {
      isAdmin = false;
    }
    if (isAdmin == false) {
      this.setState({
        TJobArr: Dep
      });
    }
    else {
      this.setState({
        TJobArr: []
      });
    }
  }

  private generateDocuments(siteUrl: string, listId: string, currentUser: string, props: IDocumentTableProps, GroupArr: any, JobArr: any, LangArr: any, siteproject: any): void {
    IService.generateDocuments(siteUrl, listId, currentUser, props, GroupArr, JobArr, LangArr, siteproject).then((response: any) => {
      items = response.results;
      this._allItems = response.results;

      if (Dep.length == 0) {
        this._allItems = [];
      }

      if (rebindDrop == true) {
        let filteredCategory = this._allItems.map((items: any) => items["Category"]);
        filteredCategory = filteredCategory.filter((x, i, a) => a.indexOf(x) == i && x != "")
        let tempArr = [];
        $.each(optionsGroup, function (e, val) {
          if (filteredCategory.indexOf(val["text"]) > -1) {
            tempArr.push(val);
          }
        })
        optionsGroup = tempArr;
        rebindDrop = false;
      }
      else {
        optionsGroup = categoryCheck;
      }

      this.setState({
        items: this._allItems,
      });



      let self = this;

      // in sequence of above jsonArray attributes values, it would be mapped one to one.
      table = $('#example').DataTable({
        scrollX: true,
        //control which datatable options available
        "info": true,
        "pagingType": 'full_numbers',
        dom: 'lBfrtip',
        buttons: [

          { extend: 'copy' },
          { extend: 'csv' },
          {
            extend: 'excel',
            text: 'Export excel',
            className: 'exportExcel',
            filename: 'Export excel',
            exportOptions: {
              modifier: {
                page: 'all'
              }
            }
          },
          {
            text: 'Json',
            action: (e, dt, node, config) => {
              var data = dt.buttons.exportData();
              var blob = new Blob([JSON.stringify(data)], { type: "text/plain;charset=utf-8" });
              FileSaver.saveAs(blob, "Document.json");
            }
          },
          { extend: 'pdf' },
          { extend: 'print' }
        ],
        data: this._allItems,
        order: [[3, 'asc']],
        columns: [
          {
            "title": "<i class='fa fa-file-o'></i>",
            render: (data, type, row) => {
              return "<input type='checkbox' class='assignPermission'/>";
            },
            "visible": isAdmin,
            orderable: false,
          },
          {
            "title": "<i class='fa fa-file-o'></i>",
            render: (data, type, row) => {
              return "<i class='fa fa-user openUserDialog'></i>";
            },
            "visible": isAdmin,
            orderable: false,
          },
          {
            "title": "<i class='fa fa-file-o'></i>",
            render: (data, type, row) => {
              return "<a><i class='fa fa-edit openDialog'></i></a>";
            },
            "visible": isAdmin,
            orderable: false,
          },
          {
            'data': 'key',
            "title": "ID"
          },
          {
            'data': 'name',
            render: (data, type, row) => {
              return "<a href='" + row["filePath"] + "' target='_blank'>" + row["name"] + "</a>";
            },
            "title": "Document Name",
          },
          {
            'data': 'DocumentType',
            "title": "DocumentType",
          },
          {
            'data': 'name',
            render: (data, type, row) => {
              return row["name"].split('-')[0];
            },
            "title": "Reference"
          },
          {
            'data': 'ReviewNo',
            "title": "ReviewNo"
          },
          {
            'data': 'Module',
            "title": "Module"
          },
          {
            'data': 'LanguageAvailable',
            "title": "LanguageAvailable"
          },
          {
            'data': 'Extension',
            "title": "Extension",
            "visible": isAdmin,
          },
          {
            'data': 'DateCreated',
            "title": "Created",
            "visible": isAdmin,
          },
          {
            'data': 'dateModified',
            "title": "Date Modifed",
            "visible": isAdmin,
          },
          {
            'data': 'FolPath',
            render: (data, type, row) => {
              return row["FilePath"];
            },
            "title": "FolPath",
            "visible": isAdmin,
          },
          {
            'data': 'Category',
            "title": "Category"
          },
          {
            'data': 'JobRole',
            "title": "JobRole",
            orderable: false,
            "visible": false
          },
          {
            'data': 'Department',
            "title": "Department",
            "visible": isAdmin,
          },
          {
            'data': 'DepartmentalOwner',
            "title": "Departmental Owner"
          },
          {
            'data': 'Groups',
            "title": "Groups",
            orderable: false,
            "visible": false
          },
          {
            'data': 'siteProjects',
            "title": "siteProjects",
            "visible": isAdmin,
          },
          {
            'data': 'DocDescription',
            "title": "Description"
          },
          {
            'data': 'Lastversiondate',
            "title": "lastversiondate",
            "visible": false
          },
          {
            'data': 'InstructionsOrNotes',
            "title": "InstructionsOrNotes"
          },
          {
            'data': 'Nextreviewdate',
            "title": "Nextreviewdate",
            "visible": isAdmin,
          },
          {
            'data': 'Applicability',
            "title": "Applicability",
            "visible": false
          },
          {
            'data': 'date',
            "title": "Date",
            "visible": isAdmin,
          },
          {
            'data': 'employeesignature',
            "title": "Employee Signature",
            "visible": isAdmin,
          },
          {
            'data': 'trainer',
            "title": "Trainer",
            "visible": true
          }
        ],
      });

      $('#example tbody').on('click', '.assignPermission', function () {
        console.log(table.row($(this).parents('tr')).data());
        console.log($(this).parents('tr').data());
        var data = table.row($(this).parents('tr')).data();
        var ischecked = $(this).is(':checked');
        self.assignPermission(data, ischecked);
      });

      $('#example tbody').on('click', '.openDialog', function () {
        var data = table.row($(this).parents('tr')).data();
        self.openProductDialog(data);
      });

      $('#example tbody').on('click', '.openUserDialog', function () {
        var data = table.row($(this).parents('tr')).data();
        self.openPermissionDialog(data);
      });

    });
  }

  public assignPermission(selectedItem: any, ischecked: any) {
    isEditDialogShow = true;
    if (ischecked == true) {
      selectDocs.push(selectedItem);
    }
    else {
      var result = selectDocs.filter((elem) => {
        return elem.key != selectedItem.key;
      });
      selectDocs = result;
    }
    console.log(selectDocs);
  }

  public openPermissionDialog(selectedItem: any) {
    itemURl = this.props.site + "/_layouts/15/user.aspx?obj=%7B218190F7-A37B-401A-9143-1B477F9DB6DE%7D," + selectedItem.key + ",LISTITEM&noredirect=true&Source=https://sjch.sharepoint.com/sites/SharedCentre/SitePages/SharedDocuments.aspx";
    isEditDialogShow = false;
    console.log(selectedItem);
    this.setState({
      hideModal: true,
      hideDialog: true,
      hidePropertyDialog: true
    });
  }

  public openProductDialog(selectedItem: any) {
    itemURl = this.props.site + "/Shared%20Documents/Forms/dispform.aspx?ID=" + selectedItem.key + "&Source=https://sjch.sharepoint.com/sites/SharedCentre/SitePages/SharedDocuments.aspx";
    isEditDialogShow = false;
    console.log(selectedItem);
    this.setState({
      hideModal: true,
      hideDialog: true,
      hidePropertyDialog: true
    });
  }

  public reloadDatatable() {
    this.componentDidMount();
  }

  public assignDocument() {
    isEditDialogShow = true;
    if (selectDocs.length > 0) {
      isDialogShow = false;
      this.setState({
        hideDialog: false,
        hidePropertyDialog: true
      });
    }
    else {
      alert("No document selected");
    }
  }

  public assignProperties() {
    isEditDialogShow = true;
    if (selectDocs.length > 0) {
      isDialogShow = false;
      this.setState({
        hideDialog: true,
        hidePropertyDialog: false
      });
    }
    else {
      alert("No document selected");
    }
  }

  public callback() {
    selectDocs = [];
    isDialogShow = true;
    this.setState({
      hideDialog: true,
      hidePropertyDialog: true
    });
    table.row().remove();
    table.clear().draw();
    table.destroy();
    this.generateDocuments(this.props.site, "Documents", this.props.currentUser, this.props, this.state.TGroupArr, this.state.TJobArr, this.state.TLangArr, this.state.TsiteProjectArr);
    //this.componentDidMount();
  }

  private _onGroupChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {

    let newSelectedItems = [...this.state.GroupArr];
    let test = [...this.state.TGroupArr];
    if (item.selected) {
      // add the option if it's checked
      newSelectedItems.push(item.key as number);
      test.push(item.text as string);
    } else {
      // remove the option if it's unchecked
      const currIndex = newSelectedItems.indexOf(item.key as number);
      const currIndex1 = test.indexOf(item.text as string);
      if (currIndex > -1) {
        newSelectedItems.splice(currIndex, 1);
      }
      if (currIndex1 > -1) {
        test.splice(currIndex, 1);
      }
    }
    this.setState({
      GroupArr: newSelectedItems,
      TGroupArr: test,
      hideModal: true,
      hideDialog: true,
      hidePropertyDialog: true
    });

    table.row().remove();
    table.clear().draw();
    table.destroy();
    this.generateDocuments(this.props.site, "Documents", this.props.currentUser, this.props, test, this.state.TJobArr, this.state.TLangArr, this.state.TsiteProjectArr);
  }

  private _onsiteProjectChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {

    let newSelectedItems = [...this.state.siteProjectArr];
    let test = [...this.state.TsiteProjectArr];
    if (item.selected) {
      // add the option if it's checked
      newSelectedItems.push(item.key as number);
      test.push(item.text as string);
    } else {
      // remove the option if it's unchecked
      const currIndex = newSelectedItems.indexOf(item.key as number);
      const currIndex1 = test.indexOf(item.text as string);
      if (currIndex > -1) {
        newSelectedItems.splice(currIndex, 1);
      }
      if (currIndex1 > -1) {
        test.splice(currIndex, 1);
      }
    }
    this.setState({
      siteProjectArr: newSelectedItems,
      TsiteProjectArr: test,
      hideModal: true,
      hideDialog: true,
      hidePropertyDialog: true
    });
    if (test.length > 0) {
      rebindDrop = true;
    } else {
      rebindDrop = false;
    }

    table.row().remove();
    table.clear().draw();
    table.destroy();
    this.generateDocuments(this.props.site, "Documents", this.props.currentUser, this.props, this.state.TGroupArr, this.state.TJobArr, this.state.TLangArr, test);
  }

  private _onRoleChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    let newSelectedItems = [...this.state.JobArr];
    let test2 = [...this.state.TJobArr];
    if (item.selected) {
      // add the option if it's checked
      newSelectedItems.push(item.key as number);
      test2.push(item.text as string);
    } else {
      // remove the option if it's unchecked
      const currIndex = newSelectedItems.indexOf(item.key as number);
      const currIndex1 = test2.indexOf(item.text as string);
      if (currIndex > -1) {
        newSelectedItems.splice(currIndex, 1);
      }
      if (currIndex1 > -1) {
        test2.splice(currIndex1, 1);
      }

    }
    if (depChecked == true) {
      $.each(Dep, function (e, val) { test2.push(val.trim()); });
    }
    this.setState({
      JobArr: newSelectedItems,
      TJobArr: test2,
      hideModal: true,
      hideDialog: true,
      hidePropertyDialog: true
    });
    console.log(JobArr);
    if (test2.length > 0) {
      rebindDrop = true;
    } else {
      rebindDrop = false;
    }

    table.row().remove();
    table.clear().draw();
    table.destroy();
    this.generateDocuments(this.props.site, "Documents", this.props.currentUser, this.props, this.state.TGroupArr, test2, this.state.TLangArr, this.state.TsiteProjectArr);

  }

  private _onJobRoleChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    //let newSelectedItems = [...this.state.jobR];
    // let test2 = [...this.state.TjobR];

    let newSelectedItems = [];
    let test2 = [];

    let temp1 = [];
    let temp2 = [];

    if (item.selected) {
      optionsJob = TempoptionsJob;
      siteprojectJob = TempsiteprojectJob;
      // add the option if it's checked
      newSelectedItems.push(item.key as number);
      test2.push(item.text as string);

      let selectedjobRole = jobRoleOption[item.text];

      let depArr = jobRoleOption[item.text].Dep.results;

      let siteArr = jobRoleOption[item.text].SiteProject.results;

      let tempJobRole = [];

      $.each(optionsJob, function (e, val) {
        if (depArr.indexOf(val.key) != -1) {
          tempJobRole.push(val);
        }
      })
      let tempSite = [];
      $.each(siteprojectJob, function (e, val) {
        if (siteArr.indexOf(val.key) != -1) {
          tempSite.push(val);
        }
      })

      $.each(tempSite, function (e, val) { temp1.push(val.text.trim()); });

      $.each(tempJobRole, function (e, val) { temp2.push(val.text.trim()); });

      this.setState({
        TsiteProjectArr: temp1,
        TJobArr: temp2,
        JobArr: depArr,
        siteProjectArr: siteArr
      })


      optionsJob = tempJobRole;
      siteprojectJob = tempSite;
    }
    else {

      // remove the option if it's unchecked
      const currIndex = newSelectedItems.indexOf(item.key as number);
      const currIndex1 = test2.indexOf(item.text as string);
      if (currIndex > -1) {
        newSelectedItems.splice(currIndex, 1);
      }
      if (currIndex1 > -1) {
        test2.splice(currIndex1, 1);
      }


      this.setState({
        TsiteProjectArr: [],
        TJobArr: [],
        JobArr: [],
        siteProjectArr: [],
        jobR: newSelectedItems,
        TjobR: test2
      })

      temp1 = [];
      temp2 = [];

      optionsJob = TempoptionsJob;
      siteprojectJob = TempsiteprojectJob;

    }
    this.setState({
      jobR: newSelectedItems,
      TjobR: test2
    })
    table.row().remove();
    table.clear().draw();
    table.destroy();
    this.generateDocuments(this.props.site, "Documents", this.props.currentUser, this.props, this.state.TGroupArr, temp2, this.state.TLangArr, temp1);

  }

  private _onLanguageChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    let newSelectedItems = [...this.state.LangArr];
    let test2 = [...this.state.TLangArr];
    if (item.selected) {
      // add the option if it's checked
      newSelectedItems.push(item.key as any);
      test2.push(item.text as string);
    } else {
      // remove the option if it's unchecked
      const currIndex = newSelectedItems.indexOf(item.key as any);
      const currIndex1 = test2.indexOf(item.text as string);
      if (currIndex > -1) {
        newSelectedItems.splice(currIndex, 1);
      }
      if (currIndex1 > -1) {
        test2.splice(currIndex, 1);
      }

    }

    this.setState({
      LangArr: newSelectedItems,
      TLangArr: test2,
    });

    table.row().remove();
    table.clear().draw();
    table.destroy();
    this.generateDocuments(this.props.site, "Documents", this.props.currentUser, this.props, this.state.TGroupArr, this.state.TJobArr, test2, this.state.TsiteProjectArr);

  }

  public dynamicSort(property) {
    var sortOrder = 1;

    if (property[0] === "-") {
      sortOrder = -1;
      property = property.substr(1);
    }

    return function (a, b) {
      if (sortOrder == -1) {
        return b[property].localeCompare(a[property]);
      } else {
        return a[property].localeCompare(b[property]);
      }
    }
  }

  public async getDropdownItems() {
    await $.ajax({
      url: "https://sjch.sharepoint.com/sites/SharedCentre/_api/web/lists/getbytitle('Document%20Categories')/items",
      method: "Get",
      headers: { "Accept": "application/json; odata=verbose" },
      success: (data) => {
        optionsGroup = [];
        categoryCheck = [];
        for (let item of data.d.results) {
          let temp = { key: item.ID, text: item.Title };
          optionsGroup.push(temp);
          categoryCheck.push(temp);
        }
        categoryCheck.sort(this.dynamicSort("text"));
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
        optionsJob = [];
        for (let item of data.d.results) {
          let temp = { key: item.ID, text: item.Code };
          optionsJob.push(temp);
          TempoptionsJob.push(temp);
        }
        optionsJob.sort(this.dynamicSort("text"));
        TempoptionsJob.sort(this.dynamicSort("text"));
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
        jobRoleOption = {};
        jobRoleOptionArr = [];
        for (let item of data.d.results) {
          let temp = { key: item.ID, text: item.Title };
          jobRoleOptionArr.push(temp);
          //jobRoleOptionArr.push(item.Title);
          jobRoleOption[item.Title] = { "Dep": item.DEPARTMENTId, "SiteProject": item.Site_x0020_LocationId }
        }
        jobRoleOptionArr.sort(this.dynamicSort("text"));
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
        siteproject = data.d.results;
        siteprojectJob = [];
        for (let item of data.d.results) {
          let temp = { key: item.ID, text: item.Title };
          siteprojectJob.push(temp);
          TempsiteprojectJob.push(temp);
        }
        siteprojectJob.sort(this.dynamicSort("text"));
        TempsiteprojectJob.sort(this.dynamicSort("text"));
      },
      error: (data, errorCode, errorMessage) => {
        alert(errorMessage);
      }
    });


    await $.ajax({
      url: "https://sjch.sharepoint.com/sites/SharedCentre/_api/web/lists/getbytitle('language')/items",
      method: "Get",
      headers: { "Accept": "application/json; odata=verbose" },
      success: (data) => {
        Langoptions = [];
        for (let item of data.d.results) {
          let temp = { key: item.ID, text: item.Title };
          Langoptions.push(temp);
        }
        Langoptions.sort(this.dynamicSort("text"));
      },
      error: (data, errorCode, errorMessage) => {
        alert(errorMessage);
      }
    });

    await $.ajax({
      url: "https://sjch.sharepoint.com/sites/SharedCentre/_api/web/lists/getbytitle('Country')/items",
      method: "Get",
      headers: { "Accept": "application/json; odata=verbose" },
      success: (data) => {
        Grpoptions = [];
        for (let item of data.d.results) {
          let temp = { key: item.ID, text: item.Title };
          Grpoptions.push(temp);
        }
        Grpoptions.sort(this.dynamicSort("text"));
      },
      error: (data, errorCode, errorMessage) => {
        alert(errorMessage);
      }
    });

  }

  public documentprint = (e) => {
    e.preventDefault();
    var divContents = document.getElementById("mypdf").innerHTML;
    var printWindow = window.open('', '', 'height=500,width=500');
    printWindow.document.write('<html><head><title>Print Page</title>');
    printWindow.document.write('<style type="text/css">');
    printWindow.document.write('@media print{.header {display: inline-block;width: 100%;}.playerOne {float: right;}.playerTwo {float: left;}#mytblpdf table{table-layout:fixed;width:500px}#mytblpdf td{border:1px solid #ddd;overflow:hidden;width:90px;word-break:break-word}#mytblpdf th{border:1px solid #ddd;overflow:hidden;width:90px;word-break:break-word}#mytblpdf th{border:1px solid #ddd;text-align:left;padding:8px;background:#03787c;color:#fff}#mytblpdf tr:nth-child(even){background-color:#ddd}#mytblpdf th{border:1px solid #ddd;text-align:left;padding:8px;background:#03787c;color:#fff}#mytblpdf td{border:1px solid #ddd;text-align:left;padding:8px}}');
    printWindow.document.write('#mytblpdf table{table-layout:fixed;width:500px}#mytblpdf td{border:1px solid #ddd;overflow:hidden;width:110px;word-break:break-word}#mytblpdf th{border:1px solid #ddd;overflow:hidden;width:110px;word-break:break-word}#mytblpdf th{border:1px solid #ddd;text-align:left;padding:8px;background:#03787c;color:#fff}#mytblpdf tr:nth-child(even){background-color:#ddd}#mytblpdf th{border:1px solid #ddd;text-align:left;padding:8px;background:#03787c;color:#fff}#mytblpdf td{border:1px solid #ddd;text-align:left;padding:8px}');
    printWindow.document.write('</style>');
    printWindow.document.write('<link rel="stylesheet" media="print" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">');
    printWindow.document.write('</head><body >');
    printWindow.document.write(divContents);
    printWindow.document.write('</body></html>');
    printWindow.document.close();
    printWindow.print();
  }

  private _showprintDialog = (): void => {

    console.log(this.state.TsiteProjectArr);
    var Temp = $.grep(siteproject, function (e, val) { return e.Title == "ALGERIA - HMD CLINIC" })

    if (Temp.length > 0) {
      var Temp1 = $.grep(Grpoptions, function (e, val) { return e.key == Temp[0].ID })
      if (Temp1.length > 0) {
        country = Temp1[0].text;
      }
    }
    this.setState({ hidePrintDialog: false });
  }

  private _closeprintDialog = (): void => {
    this.setState({ hidePrintDialog: true });
  }

  private _docNameTexthandleChange(event): void {

    this.setState({
      Text: event
    });

  }

  public _onChange(ev: React.MouseEvent<HTMLElement>, checked: boolean) {
    console.log('toggle is ' + (checked ? 'checked' : 'not checked'));
    if (checked) {
      depChecked = true;
      this.setState({
        depChecked: true,
        TJobArr: Dep,
        JobArr: [],
        LangArr: [],
        GroupArr: [],
        siteProjectArr: []
      });
      table.row().remove();
      table.clear().draw();
      table.destroy();
      this.generateDocuments(this.props.site, "Documents", this.props.currentUser, this.props, [], Dep, [], []);

    }
    else {

      depChecked = false;
      this.setState({
        depChecked: false,
        TJobArr: [],
        JobArr: [],
        LangArr: [],
        GroupArr: [],
        siteProjectArr: []
      });
      table.row().remove();
      table.clear().draw();
      table.destroy();
      this.generateDocuments(this.props.site, "Documents", this.props.currentUser, this.props, [], [], [], []);
    }

  }

  public render(): React.ReactElement<IDocumentTableProps> {

    const { GroupArr } = this.state;
    const { JobArr } = this.state;
    const { LangArr } = this.state;
    const { siteProjectArr } = this.state;
    const { jobR } = this.state;
    let today = new Date();
    let date = (today.getMonth() + 1) + '/' + today.getDate() + '/' + today.getFullYear();

    const imgStyles = {
      root: { width: 400, height: "100px" },
      image: { width: 400, height: "80px", marginTop: "5%" }
    };
    let NewObj;
    if (this.state.items) {
      NewObj = this.state.items.reduce((r, a) => {
        r[a.Module] = [...r[a.Module] || [], a];
        return r;
      }, {});
    }
    let temp;

    let HTMlString = "";

    $.each(NewObj, function (e, val) {
      temp = val;
      HTMlString += `<tr>
                    <td colspan=10 style="text-align: center;font-weight: 500;">${e}</td>
                  </tr>`;
      $.each(temp, function (e1, val1) {
        HTMlString += `<tr>
                    <td>${val1.Category}</td>
                    <td>${val1.name.split('-')[0]}</td>
                    <td>${val1.name.split('-')[1]}</td>
                    <td>${val1.DocDescription}</td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td>${val1.trainer}</td>
                  </tr>`;
      })
    });

    const DialogHTML: JSX.Element = this.state.items ?
      <div>
        <table>
          <tr>
            <th>Category</th>
            <th>Doc Reference</th>
            <th>Document Title</th>
            <th>Description & Instructions</th>
            <th>Send to</th>
            <th>When</th>
            <th>Date</th>
            <th>Employee Signature</th>
            <th>Trainer</th>
          </tr>

          {ReactHtmlParser(HTMlString)}
        </table></div>
      : <div />;

    return (
      <div>
        <table className="ms-Table" style={{ display: "none" }}>
          <thead>
            <tr>
              <th><PrimaryButton text="Assign Document" onClick={this.assignDocument} style={{ marginRight: '1%' }} /></th>
              <th> <PrimaryButton text="Assign Country and Job Role" onClick={this.assignProperties} style={{ marginRight: '1%' }} /></th>
            </tr>
          </thead>
        </table>
        <div className={depChecked == false && isAdmin == true ? styles.showbtn : styles.hidebtn}>
          <DefaultButton secondaryText="Opens the Sample Dialog" onClick={this._showprintDialog} text="Create PDF" />
        </div>

        <table className="ms-Table">
          <thead>
            <tr>
              <th className={depChecked == false && isAdmin == true ? styles.showfilter : styles.showfilter}> Filter Category (optional): <Dropdown
                placeholder="Select options"
                selectedKeys={GroupArr}
                onChange={this._onGroupChange}
                multiSelect
                options={optionsGroup}
                styles={dropdownStyles}
              /></th>
              <th className={depChecked == false && isAdmin == true ? styles.showfilter : styles.hidebtn}> Select Job Role: <Dropdown
                placeholder="Select options"
                selectedKeys={jobR}
                onChange={this._onJobRoleChange}
                multiSelect
                options={jobRoleOptionArr}
                styles={dropdownStyles}
              /></th>
              <th className={depChecked == false && isAdmin == true ? styles.showfilter : styles.hidebtn}> Select Department: <Dropdown
                placeholder="Select options"
                selectedKeys={JobArr}
                onChange={this._onRoleChange}
                multiSelect
                options={optionsJob}
                styles={dropdownStyles}
              /></th>
              <th className={depChecked == false && isAdmin == true ? styles.showfilter : styles.hidebtn}> Select Site Projects: <Dropdown
                placeholder="Select options"
                selectedKeys={siteProjectArr}
                onChange={this._onsiteProjectChange}
                multiSelect
                options={siteprojectJob}
                styles={dropdownStyles}
              /></th>
              <th className={depChecked == false && isAdmin == true ? styles.showfilter : styles.hidebtn}> Select Language: <Dropdown
                placeholder="Select options"
                selectedKeys={LangArr}
                onChange={this._onLanguageChange}
                multiSelect
                options={Langoptions}
                styles={dropdownStyles}
              /></th>
              <th className={isAdmin == true ? styles.showfilter : styles.hidebtn}>
                <Toggle label="Current User Depatment" disabled={false} style={{ minWidth: "2em" }} onChange={this._onChange} />
              </th>
              <th className={isAdmin == true ? styles.hidebtn : styles.showfilter}>
                <Toggle label="Current User Depatment" disabled={true} style={{ minWidth: "2em" }} defaultChecked onChange={this._onChange} />
              </th>
            </tr>
          </thead>
        </table>

        <table id="example" className="display"></table>

        <div>
          <Dialog
            hidden={this.state.hidePrintDialog}
            onDismiss={this._closeprintDialog}
            dialogContentProps={{
              type: DialogType.largeHeader,
              title: 'Print Document',
              subText: ''
            }}
            containerClassName={'ms-dialogMainOverride ' + styles.textDialog}
          >
            <button onClick={this.documentprint} >
              Generate Pdf  </button>



            <div id="mypdf">
              <div style={{ marginTop: "1%" }}>
                <div id="block_container" style={{ float: "left" }}>
                  <img src="https://sjch.sharepoint.com/sites/SharedCentre/PublishingImages/Medilink%20IHS%20Logo.png" alt="Trulli" width="250" height="70">
                  </img>
                </div>
              </div>
              <div className="header" style={{ marginTop: "1%" }}>
                <div className="playerOne">


                </div>
                <div className="playerTwo" >
                </div>
              </div>

              <div style={{ marginTop: "1%", marginLeft: "15%" }}>
                <div id="block_container" style={{ float: "right" }}>
                  <div style={{ fontWeight: 500, fontSize: '20px' }}>Site Projects</div>
                  <div style={{ fontWeight: 400, fontSize: '15px' }}>{this.state.TsiteProjectArr.join(',')}</div>
                  <div style={{ fontWeight: 500, fontSize: '20px' }}>Department</div>
                  <div style={{ fontWeight: 400, fontSize: '15px' }}>{this.state.TJobArr.join(',')}</div>
                  <div style={{ fontWeight: 500, fontSize: '20px' }}>Country</div>
                  <div style={{ fontWeight: 400, fontSize: '15px' }}>{country}</div>
                  <div><TextField label="Employee :" onChanged={e => this._docNameTexthandleChange(e)} value={this.state.Text} /></div>

                </div>
              </div>

              <div style={{ marginTop: "1%" }}>
                <div className="row">
                  <div className="col-sm-4" >
                    <div style={{ fontWeight: 500, fontSize: '20px' }}>
                      Induction Checklist
                    </div>
                    <div style={{ fontWeight: 500, fontSize: '16px' }}>
                      Date Document Prepared :{date}
                    </div>
                  </div>
                </div>
              </div>

              <div id="block_container">
                {/* <div style={{ fontSize: "18px", fontWeight: 400, width: '20%' }}>Employee : {this.state.Text}</div> */}
              </div>

              <div id="mytblpdf" style={{ marginTop: "3%" }}>
                {DialogHTML}
              </div>

            </div>
            <DialogFooter>
              <DefaultButton onClick={this._closeprintDialog} text="Cancel" />
            </DialogFooter>
          </Dialog>
        </div>
        <IframeDialog description="" docEditUrl={itemURl} isDlgOpen={isEditDialogShow} callback={this.reloadDatatable}></IframeDialog>
        <AssignDialog hideDialog={this.state.hideDialog} context={this.props.context} selectDocs={selectDocs} callback={this.callback}></AssignDialog>
        <AssignProperty hidePropertyDialog={this.state.hidePropertyDialog} context={this.props.context} selectDocs={selectDocs} callback={this.callback}></AssignProperty>
      </div >
    );
  }
}

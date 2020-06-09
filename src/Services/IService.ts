import { Text } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';
import * as moment from 'moment';
import pnp, { List, App, Web, CamlQuery } from "sp-pnp-js";
import { IDocumentTableProps } from '../webparts/documentTable/components/IDocumentTableProps';
import { IDocument } from '../webparts/documentTable/components/IDocumentTableState';
import * as $ from 'jquery';
export class IService {

    private spHttpClient: SPHttpClient;

    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
            this.spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
        });
    }

    public static async generateDocuments(webUrl: string, listName: string, currentUser: string, props: IDocumentTableProps, GroupsFilter: any, DepFilter: any, LangArr: any, siteproject: any): Promise<any> {
        let p = new Promise<any>(async (resolve) => {
            
            var GroupVal = "";
            $.each(GroupsFilter, (e, val) => {
                GroupVal += "<Value Type=\"LookupMulti\">" + val.toString() + "</Value>";
            });

            //var DepFilter=["ADM","MCL"];
            var DepVal = "";
            $.each(DepFilter, (e, val) => {
                DepVal += "<Value Type=\"LookupMulti\">" + val.toString() + "</Value>";
            });

            var LangVal = "";
            $.each(LangArr, (e, val) => {
                LangVal += "<Value Type=\"LookupMulti\">" + val.toString() + "</Value>";
            });

            var SiteProjectVal = "";
            $.each(siteproject, (e, val) => {
                SiteProjectVal += "<Value Type=\"LookupMulti\">" + val.toString() + "</Value>";
            });

            let caml: any = "";
            if (DepFilter.length == 0 && LangArr.length == 0 && GroupsFilter.length == 0 && siteproject.length > 0) {
                caml = "<Where>";
                caml += siteproject.length > 0 ? "<In><FieldRef Name=\"SiteProjects\" /><Values>" + SiteProjectVal + "</Values></In>" : "";
                caml += "</Where>";
            }

            else if (DepFilter.length == 0 && LangArr.length == 0 && GroupsFilter.length > 0 && siteproject.length == 0) {
                caml = "<Where>";
                caml += GroupsFilter.length > 0 ? "<In><FieldRef Name=\"Category\" /><Values>" + GroupVal + "</Values></In>" : "";
                caml += "</Where>";
            }

            else if (DepFilter.length == 0 && LangArr.length > 0 && GroupsFilter.length == 0 && siteproject.length == 0) {
                caml = "<Where>";
                caml += LangArr.length > 0 ? "<In><FieldRef Name=\"Language\" /><Values>" + LangVal + "</Values></In>" : "";
                caml += "</Where>";
            }
            else if (DepFilter.length > 0 && LangArr.length == 0 && GroupsFilter.length == 0 && siteproject.length == 0) {
                caml = "<Where>";
                caml += DepFilter.length > 0 ? "<In><FieldRef Name=\"Department\" /><Values>" + DepVal + "</Values></In>" : "";
                caml += "</Where>";
            }

            else if (DepFilter.length > 0 && LangArr.length == 0 && GroupsFilter.length == 0 && siteproject.length > 0) {
                caml = "<Where>";

                caml += DepFilter.length > 0 && siteproject.length > 0 ? "<And>" : "";
                caml += DepFilter.length > 0 ? "<In><FieldRef Name=\"Department\" /><Values>" + DepVal + "</Values></In>" : "";

                caml += siteproject.length > 0 ? "<In><FieldRef Name=\"SiteProjects\" /><Values>" + SiteProjectVal + "</Values></In>" : "";
                caml += siteproject.length > 0 && DepFilter.length > 0 ? "</And>" : "";

                caml += "</Where>";

            }
            else if (DepFilter.length > 0 && LangArr.length > 0 && GroupsFilter.length == 0 && siteproject.length == 0) {

                caml = "<Where>";

                caml += DepFilter.length > 0 && LangArr.length > 0 ? "<And>" : "";
                caml += DepFilter.length > 0 ? "<In><FieldRef Name=\"Department\" /><Values>" + DepVal + "</Values></In>" : "";

                caml += LangArr.length > 0 ? "<In><FieldRef Name=\"Language\" /><Values>" + LangVal + "</Values></In>" : "";
                caml += LangArr.length > 0 && LangArr.length > 0 ? "</And>" : "";

                caml += "</Where>";
            }
            else if (DepFilter.length > 0 && LangArr.length == 0 && GroupsFilter.length > 0 && siteproject.length == 0) {

                caml = "<Where>";

                caml += DepFilter.length > 0 && GroupsFilter.length > 0 ? "<And>" : "";
                caml += DepFilter.length > 0 ? "<In><FieldRef Name=\"Department\" /><Values>" + DepVal + "</Values></In>" : "";

                caml += GroupsFilter.length > 0 ? "<In><FieldRef Name=\"Category\" /><Values>" + GroupVal + "</Values></In>" : "";
                caml += GroupsFilter.length > 0 && DepFilter.length > 0 ? "</And>" : "";

                caml += "</Where>";

            }
            else if (DepFilter.length == 0 && LangArr.length > 0 && GroupsFilter.length > 0 && siteproject.length == 0) {
                caml = "<Where>";

                caml += LangArr.length > 0 && GroupsFilter.length > 0 ? "<And>" : "";
                caml += LangArr.length > 0 ? "<In><FieldRef Name=\"Language\" /><Values>" + LangVal + "</Values></In>" : "";

                caml += GroupsFilter.length > 0 ? "<In><FieldRef Name=\"Category\" /><Values>" + GroupVal + "</Values></In>" : "";
                caml += GroupsFilter.length > 0 && LangArr.length > 0 ? "</And>" : "";

                caml += "</Where>";
            }
            else if (DepFilter.length == 0 && LangArr.length > 0 && GroupsFilter.length == 0 && siteproject.length > 0) {
                caml = "<Where>";

                caml += LangArr.length > 0 && siteproject.length > 0 ? "<And>" : "";
                caml += LangArr.length > 0 ? "<In><FieldRef Name=\"Language\" /><Values>" + LangVal + "</Values></In>" : "";

                caml += siteproject.length > 0 ? "<In><FieldRef Name=\"SiteProjects\" /><Values>" + SiteProjectVal + "</Values></In>" : "";
                caml += siteproject.length > 0 && LangArr.length > 0 ? "</And>" : "";

                caml += "</Where>";
            }
            else if (DepFilter.length == 0 && LangArr.length == 0 && GroupsFilter.length > 0 && siteproject.length > 0) {
                caml = "<Where>";

                caml += GroupsFilter.length > 0 && siteproject.length > 0 ? "<And>" : "";
                caml += GroupsFilter.length > 0 ? "<In><FieldRef Name=\"Category\" /><Values>" + GroupVal + "</Values></In>" : "";

                caml += siteproject.length > 0 ? "<In><FieldRef Name=\"SiteProjects\" /><Values>" + SiteProjectVal + "</Values></In>" : "";
                caml += siteproject.length > 0 && LangArr.length > 0 ? "</And>" : "";

                caml += "</Where>";
            }
            else if (DepFilter.length > 0 && LangArr.length > 0 && GroupsFilter.length > 0 && siteproject.length == 0) {

                caml = "<Where>";

                caml += GroupsFilter.length > 0 && LangArr.length > 0 && DepFilter.length > 0 ? "<And><And>" : "";

                caml += GroupsFilter.length > 0 ? "<In><FieldRef Name=\"Category\" /><Values>" + GroupVal + "</Values></In>" : "";

                caml += LangArr.length > 0 ? "<In><FieldRef Name=\"Language\" /><Values>" + LangVal + "</Values></In></And>" : "";

                caml += DepFilter.length > 0 ? "<In><FieldRef Name=\"Department\" /><Values>" + DepVal + "</Values></In>" : "";

                caml += GroupsFilter.length > 0 && LangArr.length > 0 && DepFilter.length > 0 ? "</And>" : "";

                caml += "</Where>";

            }
            else if (DepFilter.length == 0 && LangArr.length > 0 && GroupsFilter.length > 0 && siteproject.length > 0) {

                caml = "<Where>";

                caml += GroupsFilter.length > 0 && LangArr.length > 0 && siteproject.length > 0 ? "<And><And>" : "";

                caml += GroupsFilter.length > 0 ? "<In><FieldRef Name=\"Category\" /><Values>" + GroupVal + "</Values></In>" : "";

                caml += LangArr.length > 0 ? "<In><FieldRef Name=\"Language\" /><Values>" + LangVal + "</Values></In></And>" : "";

                caml += siteproject.length > 0 ? "<In><FieldRef Name=\"SiteProjects\" /><Values>" + SiteProjectVal + "</Values></In>" : "";

                caml += GroupsFilter.length > 0 && LangArr.length > 0 && siteproject.length > 0 ? "</And>" : "";

                caml += "</Where>";

            }
            else if (DepFilter.length > 0 && LangArr.length == 0 && GroupsFilter.length > 0 && siteproject.length > 0) {


                caml = "<Where>";

                caml += GroupsFilter.length > 0 && DepFilter.length > 0 && siteproject.length > 0 ? "<And><And>" : "";

                caml += GroupsFilter.length > 0 ? "<In><FieldRef Name=\"Category\" /><Values>" + GroupVal + "</Values></In>" : "";

                caml += DepFilter.length > 0 ? "<In><FieldRef Name=\"Department\" /><Values>" + DepVal + "</Values></In></And>" : "";

                caml += siteproject.length > 0 ? "<In><FieldRef Name=\"SiteProjects\" /><Values>" + SiteProjectVal + "</Values></In>" : "";

                caml += GroupsFilter.length > 0 && DepFilter.length > 0 && siteproject.length > 0 ? "</And>" : "";

                caml += "</Where>";
            }
            else if (DepFilter.length > 0 && LangArr.length > 0 && GroupsFilter.length == 0 && siteproject.length > 0) {

                caml = "<Where>";

                caml += LangArr.length > 0 && DepFilter.length > 0 && siteproject.length > 0 ? "<And><And>" : "";

                caml += LangArr.length > 0 ? "<In><FieldRef Name=\"Language\" /><Values>" + LangVal + "</Values></In>" : "";

                caml += DepFilter.length > 0 ? "<In><FieldRef Name=\"Department\" /><Values>" + DepVal + "</Values></In></And>" : "";

                caml += siteproject.length > 0 ? "<In><FieldRef Name=\"SiteProjects\" /><Values>" + SiteProjectVal + "</Values></In>" : "";

                caml += LangArr.length > 0 && DepFilter.length > 0 && siteproject.length > 0 ? "</And>" : "";

                caml += "</Where>";

            }
            else if (DepFilter.length > 0 && LangArr.length > 0 && GroupsFilter.length > 0 && siteproject.length > 0) {

                caml = "<Where>";

                caml += LangArr.length > 0 && DepFilter.length > 0 && siteproject.length > 0 && GroupsFilter.length > 0 ? "<And><And>" : "";

                caml += LangArr.length > 0 ? "<In><FieldRef Name=\"Language\" /><Values>" + LangVal + "</Values></In>" : "";

                caml += DepFilter.length > 0 ? "<In><FieldRef Name=\"Department\" /><Values>" + DepVal + "</Values></In></And><And>" : "";

                caml += siteproject.length > 0 ? "<In><FieldRef Name=\"SiteProjects\" /><Values>" + SiteProjectVal + "</Values></In>" : "";

                caml += GroupsFilter.length > 0 ? "<In><FieldRef Name=\"Category\" /><Values>" + GroupVal + "</Values></In></And>" : "";

                caml += LangArr.length > 0 && DepFilter.length > 0 && siteproject.length > 0 && GroupsFilter.length > 0 ? "</And>" : "";
                
                caml += "</Where>";
            }

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
                        <FieldRef Name="Created"/>
                        <FieldRef Name="Modified"/>
                        <FieldRef Name="docIcon"/>
                        <FieldRef Name="DocTitle"/>
                        <FieldRef Name="Modified"/>

                        <FieldRef Name="Reference"/>	
                        <FieldRef Name="ReviewNo"/>
                        <FieldRef Name="Module"/>

                        <FieldRef Name="DocumentType"/>
                        <FieldRef Name="Language"/>
                        <FieldRef Name="Extension"/>
                        <FieldRef Name="DocDescription"/>

                        <FieldRef Name="DateCreated"/>
                        <FieldRef Name="Lastversiondate"/>
                        <FieldRef Name="Department"/>
                        <FieldRef Name="Category"/>

                        <FieldRef Name="InstructionsOrNotes"/>

                        <FieldRef Name="Nextreviewdate"/>	
                        <FieldRef Name="JobRole"/>

                        <FieldRef Name="Applicability"/>	
                        <FieldRef Name="Groups"/>
                        <FieldRef Name="SiteProjects"/>
                        <FieldRef Name="Departmental_x0020_Owner"/>
                        <FieldRef Name="Trainer"/>
                        <FieldRef Name="SAVE ON SERVER"/>
                        <FieldRef Name="COMPILED BY"/>
                        <FieldRef Name="Send to"/>
                        <FieldRef Name="When"/>
                        <FieldRef Name="Doc Review Date"/>                        
                        <FieldRef Name="Doc Reminder Date for Review"/>
                      

                    </ViewFields>`;

            const countQuery: CamlQuery = {
                ViewXml: `${camlQuery}</View>`,
            };

            let response = this.BindWorkItems(webUrl, listName, countQuery);
            resolve(response);
        });
        return p;
    }


    /***************Bind Ideas************/
    public static async BindWorkItems(siteUrl: string, listName: string, countQuery: CamlQuery) {
        let web = new Web(siteUrl);
        const result = await web.lists.getByTitle(listName).getItemsByCAMLQuery(countQuery, 'FieldValuesAsText', "FileLeafRef", "FileRef");
        var response: any = {};
        let ideas: IDocument[] = [];
        let fileURL;

        for (const item of result) {
            var fileType = item.FileLeafRef.split('.');
            var ext = fileType.slice(fileType.length - 1, 5);
            if (ext == "dot" || ext == "pot" || ext == "doc") {
                fileURL = "https://spoprod-a.akamaihd.net/files/fabric/assets/item-types-fluent/20/" + ext + "x.svg";
            }
            else {
                fileURL = "https://spoprod-a.akamaihd.net/files/fabric/assets/item-types-fluent/20/" + ext + ".svg";
            }
            let siteURLArr = siteUrl.split('/');
            let webURL = siteURLArr.splice(0, 3).join("/");
            //  this._generateDocuments();
            let editUser = siteUrl + "/_layouts/15/user.aspx?obj=%7B218190F7-A37B-401A-9143-1B477F9DB6DE%7D," + item.Id + ",LISTITEM&noredirect=true";
            let fileDocURL = siteUrl + "/_layouts/15/WopiFrame.aspx?sourcedoc=" + encodeURIComponent(item.FileRef) + "&action=default";
            ideas.push({
                dateModified: item.FieldValuesAsText.Modified,
                iconName: fileURL,
                filePath: fileDocURL,
                key: item.ID,
                name: item.FileLeafRef,
                FilePath: item.FileRef,
                userEdit: editUser,
                Reference: item.Reference != null ? item.Reference : "",
                ReviewNo: item.ReviewNo != null ? item.ReviewNo : "",
                Module: item.FieldValuesAsText.Module,
                DocumentType: item.FieldValuesAsText.DocumentType,
                LanguageAvailable: item.FieldValuesAsText.Language,
                Extension: item.FileLeafRef.split('.')[1],
                DocDescription: item.DocDescription != null ? item.DocDescription : "",
                DateCreated: item.FieldValuesAsText.Created,
                Lastversiondate: item.FieldValuesAsText.Lastversiondate,
                Department: item.DepartmentId.length > 0 ? item.FieldValuesAsText.Department : "",
                Category: item.CategoryId.length > 0 ? item.FieldValuesAsText.Category : "",
                InstructionsOrNotes: item.InstructionsOrNotes != null ? item.InstructionsOrNotes : "",
                Nextreviewdate: item.FieldValuesAsText.Nextreviewdate,
                JobRole: item.FieldValuesAsText.JobRole,
                Groups: item.GroupsId.length > 0 ? item.FieldValuesAsText.Groups : "",
                siteProjects: item.SiteProjectsId.length > 0 ? item.FieldValuesAsText.SiteProjects : "",
                Applicability: item.FieldValuesAsText.Applicability,
                date: "",
                employeesignature: "",
                trainer: item.FieldValuesAsText.Trainer,
                Trainer: item.FieldValuesAsText.Trainer,
                DocReminder: "",
                Departmental: "",
                DocReview: "",
                DepartmentalOwner: item.FieldValuesAsText.Departmental_x005f_x0020_x005f_Owner,
                When: item.FieldValuesAsText.When,
                Sendto: "",
                COMPILED: "",
                SAVE: "",
            });
        }

        response.results = ideas;
        return response;
    }
}

import { IColumn } from "office-ui-fabric-react/lib/DetailsList";

export interface IDocumentTableState {
    // columns: IColumn[];
    items: IDocument[];
    //selectionDetails: string;
    isModalSelection: boolean;
    isCompactMode: boolean;
    announcedMessage?: string;
    hideDialog: boolean;
    hidePropertyDialog: boolean;
    hideModal: boolean;
    hidePrintDialog: boolean;
    GroupItem?: number[];
    JobRoleItem?: number[];
    GroupArr: number[];
    siteProjectArr: number[];
    TsiteProjectArr: string[];
    JobArr: number[];
    TGroupArr: string[];
    LangArr: string[];
    TLangArr: string[];
    TJobArr: string[];
    TJobArrNew: string[];
    Text: string;
    depChecked: boolean;

    jobR: number[];
    TjobR: string[];
}

export interface IDocument {
    dateModified: string;
    iconName: string;
    key: string;
    name: string;
    FilePath: string;
    Reference: string;
    ReviewNo: string;
    Module: string;
    DocumentType: string;
    LanguageAvailable: string;
    Extension: string;
    DocDescription: string;
    DateCreated: string;
    Lastversiondate: string;
    Department: string;
    Category: string;
    InstructionsOrNotes: string;
    Nextreviewdate: string;
    JobRole: string;
    Groups: string;
    siteProjects: string;
    Applicability: string;
    filePath: string;
    userEdit: string;
    date: string;
    employeesignature: string;
    trainer: string;
    Trainer: string;
    DocReminder: string;
    Departmental: string;
    DocReview: string;
    DepartmentalOwner: string;
    When: string;
    Sendto: string;
    COMPILED: string;
    SAVE: string;

}
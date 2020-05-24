import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IAssignPropertyDialogProps {
    hidePropertyDialog: boolean;
    context: WebPartContext;
    selectDocs:any[];
    callback:Function;
   }
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IAssignDialogProps {
    hideDialog: boolean;
    context: WebPartContext;
    selectDocs:any[];
    callback:Function;
   }
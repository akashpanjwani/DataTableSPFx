import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { IDocumentTableState } from './IDocumentTableState';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ServiceScope } from '@microsoft/sp-core-library';


export interface IDocumentTableProps {
  List: any;
  WidgetChoice: string;
  title: string;
  site: string;
  currentUser: string;
  serviceScope: ServiceScope;
  numberOfWorkItemsToShow: string;
  context: WebPartContext;
}

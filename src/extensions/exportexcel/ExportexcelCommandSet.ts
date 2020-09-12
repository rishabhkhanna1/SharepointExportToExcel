import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters,
  RowAccessor
} from '@microsoft/sp-listview-extensibility';
// import { Dialog } from '@microsoft/sp-dialog';
import * as xlsx from 'xlsx';
import { SPHttpClient,SPHttpClientResponse} from '@microsoft/sp-http';
// import * as strings from 'ExportexcelCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IExportexcelCommandSetProperties {

}

const LOG_SOURCE: string = 'ExportexcelCommandSet';

export default class ExportexcelCommandSet extends BaseListViewCommandSet<IExportexcelCommandSetProperties> 
{

  
  private _wb;
  private _viewColumns: string[];
  private _listTitle: string;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ExportexcelCommandSet');
    this.Initiate();
    return Promise.resolve();
  }

  private async Initiate() {
    await this.getViewColumns();
  }

  private async getViewColumns(){
    const currentWebUrl: string = this.context.pageContext.web.absoluteUrl;
    this._listTitle = this.context.pageContext.legacyPageContext.listTitle;
    const viewId: string = this.context.pageContext.legacyPageContext.viewId.replace('{','').replace('}','');
    this.context.spHttpClient.get(`${currentWebUrl}/_api/lists/getbytitle('${this._listTitle}')/Views('${viewId}')/ViewFields`,SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
        response.json().then((viewColumnsResponse: any) => {          
          this._viewColumns = viewColumnsResponse.Items;
        });
      });
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const exportCommand: Command = this.tryGetCommand('EXCELEXPORTITEMS_1');
    if (exportCommand) {
      exportCommand.visible = event.selectedRows.length > 0;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void 
  {
    let _grid: any[];
    let index=this._viewColumns.indexOf('LinkTitle');
    if(index !== -1)
    {
      this._viewColumns[index]='Title';
    }

    switch (event.itemId)
    {
      case 'EXCELEXPORTITEMS_1':
        if(event.selectedRows.length > 0)
        {
          _grid= new Array(event.selectedRows.length);
          _grid[0]=this._viewColumns;

          event.selectedRows.forEach((row: RowAccessor,index: number) =>{
            let _row: string[] = [];
            let i: number=0;
            this._viewColumns.forEach((viewColumn: string) => {
              _row[i++]=this._getFieldValuesAsText(row.getValueByName(viewColumn));
            });
            _grid[index+1] = _row;
          });
        }
        break;
     
      default:
        throw new Error('Unknown command');
    }
    this.writeToExcel(_grid);
  }
  private writeToExcel(data: any[]): void
    {
    let ws = xlsx.utils.aoa_to_sheet(data);
    let wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, 'selected-items');
    xlsx.writeFile(wb, `${this._listTitle}.xlsx`);
    }

    private _getFieldValuesAsText(field: any) : string 
    {

      let fieldValue: string;
      switch(typeof field)
      {
        case 'object':{
          if(field instanceof Array)
          {
            if(!field.length)
            {
              fieldValue='';
            }
            else if(field[0].title)
            {
              fieldValue=field.map(value => value.title).join();
            }
            else if(field[0].lookupValue)
            {
              fieldValue=field.map(value => value.lookupValue).join();
            }
            else if(field[0].Label)
            {
              fieldValue=field.map(value => value.Label).join();
            }
          }
          break;
        }
        default:
          {
            fieldValue=field;
          }
      }
      return fieldValue;
    }
}


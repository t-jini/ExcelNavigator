import { actiondone,HostUpdate_SetWorkbook,HostUpdate_SetTablesAndCharts,
    HostUpdate_WorksheetSelectionchanged,HostUpdate_WorksheetDatachanged,HostUpdate_WorksheetDeleted, HostUpdate_WorksheetAdded, HostUpdate_WorksheetActivated, 
    HostUpdate_WorksheetDeactivated } from '../actions/Excel'
import {} from '../actions/ExcelAction'
class ExcelAdapterClass{
    constructor (){
        this.store;
    }
    setStore(store){
        this.store = store;
    }
    errorHandler = (error)=>{
        console.log("Error: " + error);

        if (error instanceof window.OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }
    onSheetDeleted = (args)=>{
        let worksheetId = args.name;
        this.store.dispatch(HostUpdate_WorksheetDeleted(worksheetId));
    }

    onSheetAdded = (args) =>{
        let store = this.store;
        window.Excel.run(function (ctx) {
            let worksheetId = args.worksheetId; 
            let worksheet = ctx.workbook.worksheets.getItem(worksheetId);
            worksheet.load("name,id");
            return ctx.sync().then(()=>{
                store.dispatch(HostUpdate_WorksheetAdded({name : worksheet.name,id : worksheet.id}));
                store.dispatch(HostUpdate_WorksheetActivated({name : worksheet.name,id : worksheet.id}));
            });
        });
    }

    onSheetDeactivated = (args) => {
        let worksheet = args.worksheet;
        worksheet.load("name,id");
        return worksheet.context.sync().then(()=>{
            this.store.dispatch(HostUpdate_WorksheetDeactivated({name : worksheet.name, id : worksheet.id}));
        });
    }

    onSheetSelectionChanged = (args) =>{
        let store = this.store;
        window.Excel.run(function (ctx) {
            let worksheetId = args.worksheetId;
            let address = args.address;
            let worksheet = ctx.workbook.worksheets.getItem(worksheetId);
            worksheet.load("name,id");
            return ctx.sync().then(()=>{
                store.dispatch(HostUpdate_WorksheetSelectionchanged({worksheet :{name : worksheet.name, id : worksheet.id},range : address}));
            });
        });
    }

    onSheetDataChanged = (args) => {
        let address = args.address;
        let worksheet = args.worksheet;
        let changeType = args.changeType; 
        let type = args.type;
        let source = args.source;
        let store = this.store;
        window.Excel.run(function (ctx) {
            let worksheetId = args.worksheetId;
            let worksheet = ctx.workbook.worksheets.getItem(worksheetId);
            worksheet.load("name,id");
            return ctx.sync().then(()=> {
                store.dispatch(HostUpdate_WorksheetDatachanged({worksheet :{name : worksheet.name, id : worksheet.id},datachangedstate : {isChanged : true,changeType,range:address, source}}));
                //this.store.dispatch(HostUpdate_WorksheetSelectionchanged({worksheet :{name : worksheet.name, id : worksheet.id},range : address}));
            });
        });        

    }
    onSheetActivated = (args) =>{
        let store = this.store;
        window.Excel.run(function (ctx) {
            let worksheetId = args.worksheetId;
            let worksheet = ctx.workbook.worksheets.getItem(worksheetId);
            worksheet.load("name,id");
            return ctx.sync().then(()=>{
                store.dispatch(HostUpdate_WorksheetActivated({name : worksheet.name, id : worksheet.id}));
            });
        });
    }
    Adapter_ListenWorkbook = (action,state) => {
        window.Excel.run(function (ctx) {
            let worksheets = ctx.workbook.worksheets;
            //worksheets.onDeleted.add(ExcelAdapter.onSheetDeleted);
            worksheets.onAdded.add(ExcelAdapter.onSheetAdded);
            return ctx.sync().then(()=>ExcelAdapter.store.dispatch(actiondone(state.nums + 1)));
        }).catch(this.errorHandler);
    }           
    Adapter_ListenWorksheet = (action,state) =>{
        window.Excel.run(function (ctx) {
            let worksheet = ctx.workbook.worksheets.getItem(action.state.worksheetId);
            return ctx.sync().then(()=>{ 
                worksheet.onActivated.add(ExcelAdapter.onSheetActivated);
                //worksheet.onDeactivated.add(this.onSheetDeactivated);
                worksheet.onSelectionChanged.add(ExcelAdapter.onSheetSelectionChanged);
                worksheet.onDataChanged.add(ExcelAdapter.onSheetDataChanged);
                    return ctx.sync().then(()=>ExcelAdapter.store.dispatch(actiondone(state.nums + 1)));
            }).catch(this.errorHandler);
        }).catch(this.errorHandler);
    }

    Adapter_AddTablesANDChartes = (action,state)=>{
            window.Excel.run(function (ctx) {
            let worksheet = ctx.workbook.worksheets.getItem(action.state.worksheetId);
            worksheet = worksheet.load("id,name")
            let charts = worksheet.charts.load('item');
            let tables = worksheet.tables.load('item');
            
            
            return ctx.sync().then(function () {
                let state = {id:worksheet.id,name:worksheet.name,tables : [], charts : []};
                let ranges = [];
                for (var j = 0; j < tables.items.length; j++) {
                    console.log(tables.items[j]);
                    let range = tables.items[j].getRange().load('address');
                    
                    ranges.push(range);
                    
                    tables.items[j].load('name,id');
                }
                for (var j = 0; j < charts.items.length; j++) {
                    charts.items[j].load('name,id');
                }
                
                return ctx.sync().then(function () {
                    for (var j = 0; j < tables.items.length; j++) {
                        state.tables.push({
                            name : tables.items[j].name,
                            id : tables.items[j].id,
                            range : ranges[j].address
                        });
                    }
                    for (var j = 0; j < charts.items.length; j++) {
                        state.charts.push({
                            name : charts.items[j].name,
                            id : charts.items[j].id
                        });
                    }
                    ExcelAdapter.store.dispatch(HostUpdate_SetTablesAndCharts(state));
                    ExcelAdapter.store.dispatch(actiondone(state.nums + 1));
                }).catch(this.errorHandler);
            }).catch(this.errorHandler);
            }).catch(this.errorHandler);
    }
    Adapter_DeleteWorksheet = (action,state) =>{
        let worksheetId = action.state.worksheetId;
        
        window.Excel.run(function (ctx) {
            var worksheet = ctx.workbook.worksheets.getItem(worksheetId);
            worksheet.delete();
            return ctx.sync().then(()=>ExcelAdapter.store.dispatch(actiondone(state.nums + 1)));
        }).catch(this.errorHandler);
    }
    Adapter_ActivateWorksheet = (action,state) =>{
        let worksheetId = action.state.worksheetId;
        window.Excel.run(function (ctx) {
            var worksheet = ctx.workbook.worksheets.getItem(worksheetId);
            worksheet.activate();
            return ctx.sync();
        }).catch(this.errorHandler);
    }
    Adapter_ColorRange = (action,state) =>{
        window.Excel.run(function (ctx) {
            let worksheet = ctx.workbook.worksheets.getItem(action.state.worksheetId).load("name,id");
            let sourceRange = worksheet.getRange(action.state.rangeaddress);
            sourceRange.load("format");
            return ctx.sync().then(function () {
                sourceRange.format.fill.color = action.state.color;
                return ctx.sync().then(()=>ExcelAdapter.store.dispatch(actiondone(state.nums + 1)));
            });
        });
    }    

    Adapter_LoadWorkbook = (action,state)=>{
        
        window.Excel.run(function (ctx) {
            let state = {};
            let activesheet = ctx.workbook.worksheets.getActiveWorksheet().load('name,id');
            let worksheets = ctx.workbook.worksheets;
            worksheets.load('items');
            
            return ctx.sync().then(function () {
                state.activesheet = {id:activesheet.id,name:activesheet.name};
                state.worksheets = [];
                for (let i = 0; i < worksheets.items.length; i++) {
                    worksheets.items[i].load('name,id');
                }
                return ctx.sync().then(function () {
                    let worksheets = ctx.workbook.worksheets;
                    for (let i = 0; i < worksheets.items.length; i++) {
                        state.worksheets.push({id:worksheets.items[i].id,name:worksheets.items[i].name,tables : [], charts : []});
                    }
                    
                    ExcelAdapter.store.dispatch(HostUpdate_SetWorkbook(state));
                    ExcelAdapter.store.dispatch(actiondone(state.nums + 1));
                    
                }).catch(this.errorHandler);
            }).catch(this.errorHandler);
        });
    }
}
 export const ExcelAdapter = new ExcelAdapterClass();
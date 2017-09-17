import React from 'react'
import { connect } from 'react-redux'
import { displaytips, displaypaint } from '../actions'
import { bindActionCreators } from 'redux';
class ExcelModule extends React.Component{
    constructor(props){
        super(props); 
        /*this.onSheetDeleted = this.onSheetDeleted.bind(this);
        this.onSheetAdded = this.onSheetAdded.bind(this);
        this.onSheetActivated = this.onSheetActivated.bind(this);
        this.onSheetSelectionChanged = this.onSheetSelectionChanged.bind(this);
        this.onSheetDataChanged = this.onSheetDataChanged.bind(this);*/
    }
    /*onSheetDeleted(args){
        let worksheetId = args.name;
        this.props.actions.DeleteWorksheet(worksheetId);
    }

    onSheetAdded(args){
        let worksheet = args.worksheet;
        worksheet.load("name,id");
        let props = this.props;
        return worksheet.context.sync().then(()=>{
            props.actions.AddWorksheet({name : worksheet.name,id : worksheet.id});
            props.actions.ActivateWorksheet({name : worksheet.name,id : worksheet.id});
        });
    }

    

    onSheetDeactivated(args){
        let worksheet = args.worksheet;
        worksheet.load("name,id");
        let props = this.props;
        return worksheet.context.sync().then(()=>{
            props.actions.DeactivateWorksheet({name : worksheet.name, id : worksheet.id});
        });
    }

    onSheetSelectionChanged(args){
        console.log(args);
        let address = args.address;
        let worksheet = args.worksheet;
        worksheet.load("name,id");
        let props = this.props;
        return args.context.sync().then(()=>{
            
            props.actions.SelectionchangedWorksheet({worksheet :{name : worksheet.name, id : worksheet.id},range : address});
            //props.actions.DatachangedWorksheet({worksheet :{name : worksheet.name, id : worksheet.id},datachangedstate : {isChanged : true,changeType : "asd",range}});
        });
    }

    onSheetDataChanged(args){
        let address = args.address;
        let worksheet = args.worksheet;
        let changeType = args.type;
        let props = this.props;
        worksheet.load("name,id");
        return worksheet.context.sync().then(()=> {
            props.actions.DatachangedWorksheet({worksheet :{name : worksheet.name, id : worksheet.id},datachangedstate : {isChanged : true,changeType,range:address}});
            props.actions.SelectionchangedWorksheet({worksheet :{name : worksheet.name, id : worksheet.id},range : address});
        });
    }*/
    componentWillUpdate(nextProps, nextState){
    
        /*let errorHandler = this.errorHandler;
        for (let i = 0; i < nextProps.config.actions.length; i++){
            let action = nextProps.config.actions[i];
            if (action.done) continue;
            this.props.actions.actiondone(action.id);
            console.log(action.type + " " + action.id);
            switch (action.type) {
                case 'LISTEN_WORKBOOK':{
                    window.Excel.run(function (ctx) {
                        let worksheets = ctx.workbook.worksheets;
                        //worksheets.onDeleted.add(this.onSheetDeleted);
                        worksheets.onAdded.add(this.onSheetAdded);
                        return ctx.sync();
                    }.bind(this)).catch(errorHandler);
                    break;
                }           
                case 'LISTEN_WORKSHEET':
                {
                    window.Excel.run(function (ctx) {
                        let worksheet = ctx.workbook.worksheets.getItem(action.state.worksheetId);
                        return ctx.sync().then(()=>{ 
                            worksheet.onActivated.add(onSheetActivated);
                            //worksheet.onDeactivated.add(this.onSheetDeactivated);
                            worksheet.onSelectionChanged.add(this.onSheetSelectionChanged);
                            worksheet.onDataChanged.add(this.onSheetDataChanged);
                             return ctx.sync();
                        }).catch(errorHandler);
                    }.bind(this)).catch(errorHandler);
                    break;
                }

                case 'ADD_TABLESANDCHARTS':
                {
                    let props = this.props;
                    
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
                                props.actions.SetTablesAndCharts(state);
                            }).catch(errorHandler);
                        }).catch(errorHandler);
                     }.bind(this)).catch(errorHandler);
                    break;
                }
                case 'OP_DELETE_WORKSHEET':{
                    let worksheetId = action.state.worksheetId;
                    
                    window.Excel.run(function (ctx) {
                        var worksheet = ctx.workbook.worksheets.getItem(worksheetId);
                        worksheet.delete();
                        return ctx.sync();
                    }).catch(errorHandler);
                    break;
                }
                case 'OP_ACTIVATE_WORKSHEET':{
                    let worksheetId = action.state.worksheetId;
                    window.Excel.run(function (ctx) {
                        var worksheet = ctx.workbook.worksheets.getItem(worksheetId);
                        worksheet.activate();
                        return ctx.sync();
                    }).catch(errorHandler);
                    break;
                }
                case 'OP_COLORED_RANGE':{
                    window.Excel.run(function (ctx) {
                        let worksheet = ctx.workbook.worksheets.getItem(action.state.worksheetId).load("name,id");
                        let sourceRange = worksheet.getRange(action.state.rangeaddress);
                        sourceRange.load("format");
                        return ctx.sync().then(function () {
                            sourceRange.format.fill.color = action.state.color;
                            return ctx.sync();
                        });
                    });
                    break;
                }
                
            }
        }*/
    }
    errorHandler(error){
        console.log("Error: " + error);

        if (error instanceof window.OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }
    render(){
        return (<div></div>)
    }
  
}

function mapStateToProps(state) {
  const props = { workbook : state.workbook, config : state.config };
  return props;
}
function mapDispatchToProps(dispatch) {
  const actions = {};
  const actionMap = { actions: bindActionCreators(actions, dispatch) };
  //console.log(actionMap.actions.addTodo);
  return actionMap;
}
ExcelModule = connect(mapStateToProps, mapDispatchToProps)(ExcelModule)

export default ExcelModule

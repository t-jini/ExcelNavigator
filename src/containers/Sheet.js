import React from 'react'
import { connect } from 'react-redux'
import { displaytips } from '../actions'
import { ColoredRange,ActivateWorksheet,DeleteWorksheet,ListenWorksheet,AddTablesAndCharts} from '../actions/ExcelAction'
import { HostUpdate_WorksheetDatachanged } from '../actions/Excel'
import { bindActionCreators } from 'redux'
import { Collapse } from 'react-bootstrap'
import TableList from './TableList'
class Sheet extends React.Component{
    constructor(props){
        super(props);
        this.activeSheet = this.activeSheet.bind(this);
        this.collapse = this.collapse.bind(this);
        this.delete = this.delete.bind(this);
        this.state = {open : false};
    }

  
    componentWillUnmount(){
        let props = this.props;
        let text = <div>Delete : {props.worksheet.name} <br/> </div>
            
        props.actions.displaytips({text,active:true});
    }
    componentWillMount(){
        let props = this.props;
        
        let text = "WorkSheet :"+this.props.worksheet.name+" is Added."
        this.props.actions.displaytips({text,active:true});
        
        this.props.actions.ListenWorksheet({worksheetId : this.props.worksheet.id});
        this.props.actions.AddTablesAndCharts({worksheetId :this.props.worksheet.id});
 
    }

    componentWillUpdate(nextProps, nextState){
        let props = this.props;
        if ((nextProps.workbook.activesheet.name == nextProps.worksheet.name) && (this.props.workbook.activesheet.name != this.props.worksheet.name)){
            let text = <div>WorkSheet :{this.props.worksheet.name}<br/> is activated.</div>
            this.props.actions.displaytips({text,active:true});
        }
        if (nextProps.worksheet.datachangedstate && nextProps.worksheet.datachangedstate.isChanged){
            console.log(this.props.worksheet.datachangedstate);
            let changeStatus = nextProps.worksheet.datachangedstate;
            let text = <div>Type : {changeStatus.type} <br/> ChangedType : {changeStatus.changeType} <br/> RangeAddress : {changeStatus.range} <br/> Source : {changeStatus.source} </div>
            props.actions.HostUpdate_WorksheetDatachanged({worksheet :{name : this.props.worksheet.name, id : this.props.worksheet.id},datachangedstate :{isChanged : false}});
            props.actions.displaytips({text,active:true});
            let worksheet = nextProps.worksheet;
            //props.actions.datachangedworksheet({worksheet :{name : worksheet.name, id : worksheet.id},datachangedstate :{isChanged : false}});
            /*if (!this.timeOutTips)
            this.timeOutTips = setTimeout(()=>{
                props.actions.HostUpdate_WorksheetDatachanged({worksheet :{name : this.props.worksheet.name, id : this.props.worksheet.id},datachangedstate :{isChanged : false}});
                this.timeOutTips = undefined;
            },2000);*/
        }
        
        if (props.display.paint && (nextProps.worksheet.selectionaddress != this.props.worksheet.selectionaddress)){
            if (nextProps.worksheet.tables){
                window.Excel.run(function (ctx) {
                    let ranges = [];
                    console.log(nextProps.worksheet.tables.length);
                    for (let i = 0; i < nextProps.worksheet.tables.length; i++){
                        let range = ctx.workbook.worksheets.getItem(nextProps.worksheet.id).getRange(nextProps.worksheet.tables[i].address).getIntersection(nextProps.worksheet.selectionaddress).load("address");
                        ranges.push({id:nextProps.worksheet.tables[i].id,range});
                    }
                    return ctx.sync().then(()=>{
                        for (let i = 0; i < ranges.length; i++){
                            
                            if (ranges[i].range.address){
                                console.log(props.workbook);
                                props.actions.HostUpdate_ActivateTable({tableid:ranges[i].id});
                            }
                        }
                    });
                }.bind(this));
            }
            
            this.props.actions.ColoredRange({worksheetId : this.props.worksheet.id,color : "red", rangeaddress : nextProps.worksheet.selectionaddress});
        }
    }

    collapse(){
        this.setState({ open: !this.state.open });
    }

    activeSheet(){
        this.props.actions.ActivateWorksheet({worksheetId : this.props.worksheet.id});
    }
    delete(){
        this.props.actions.DeleteWorksheet({worksheetId : this.props.worksheet.id});
    }
    render(){
        let arrow = "";
        if (this.props.worksheet.tables.length > 0 || 
            this.props.worksheet.charts.length > 0) {
                arrow =  <span className="arrow" onClick = {()=>this.collapse()}></span>
        }
        return (
            <div >
                <li className = {`
                    ${this.props.workbook.activesheet.name == this.props.worksheet.name ? 'active' : ''} 
                    ${this.props.worksheet.datachangedstate && this.props.worksheet.datachangedstate.isChanged ? 'flicker' : ''}
                `}
                onClick = {() => this.activeSheet()}>
                    <a className="sheetname" href="#">
                        
                        <i className="ms-Icon ms-Icon--ExcelDocument" style={{fontSize : "15px"}}></i>
                        
                        {this.props.worksheet.name}
                        {this.props.worksheet.ischanged ? "isChanged" : ""}
                        {arrow}
                    </a>
                    <span className = "removebtn">
                        <i className="ms-Icon ms-Icon--Cancel" onClick = {() => this.delete()}></i>
                    </span>
                </li>
                
                <Collapse in = {this.state.open}>
                    <div>
                        <TableList tables = {this.props.worksheet.tables}/>
                    </div>
                </Collapse>
            </div>            
        );
    }
}
function mapStateToProps(state) {
  const props = { workbook : state.workbook,display : state.display };
  return props;
}
function mapDispatchToProps(dispatch) {
  const actions = {ColoredRange,ActivateWorksheet,DeleteWorksheet,displaytips,ListenWorksheet,AddTablesAndCharts,HostUpdate_WorksheetDatachanged};
  const actionMap = { actions: bindActionCreators(actions, dispatch) };
  //console.log(actionMap.actions.addTodo);
  return actionMap;
}
Sheet = connect(mapStateToProps, mapDispatchToProps)(Sheet)

export default Sheet

import React from 'react'
import { connect } from 'react-redux'
import { HostUpdate_ActivateTable } from '../actions/Excel'
import { bindActionCreators } from 'redux'
class Table extends React.Component{
    constructor(props){
        super(props);
        this.activeSheet = this.activeTable.bind(this);
        //console.log(props);
    }
    
    errorHandler(error){
        console.log("Error: " + error);
        if (error instanceof window.OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }
    activeTable(){
        console.log(this.props.workbook.activetableid );
        this.props.actions.HostUpdate_ActivateTable({tableid:this.props.table.id});
    }
    render(){
        //for (let i = 0; i < workbook)
        return (
            <li className = {this.props.workbook.activetableid == this.props.table.id  ? "active" : ""} onClick = {() => this.activeTable()}>
                <a href="#">
                    <i className="ms-Icon ms-Icon--Table" style={{fontSize : "15px"}}></i>
                    {this.props.table.name}
                </a>
            </li>         
        );
    }
}
function mapStateToProps(state) {
  const props = { workbook : state.workbook };
  return props;
}
function mapDispatchToProps(dispatch) {
  const actions = {HostUpdate_ActivateTable};
  const actionMap = { actions: bindActionCreators(actions, dispatch) };
  //console.log(actionMap.actions.addTodo);
  return actionMap;
}
Table = connect(mapStateToProps, mapDispatchToProps)(Table)

export default Table

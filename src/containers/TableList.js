import React from 'react'
import { connect } from 'react-redux'
import { bindActionCreators } from 'redux';
import  Table  from './Table'
class TableList extends React.Component{
    constructor(props){
        super(props);
        
        /*window.Excel.run(function (ctx){
            let sheet = ctx.workbook.worksheets.getActiveWorksheet();
            sheet.load("name");
            return ctx.sync().then(function(){
                dispatch(addTodo("iweuryhi"));
            })
        });*/
    }
    
    errorHandler(error){
        console.log("Error: " + error);
        if (error instanceof window.OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }
    render(){
        if (this.props.tables != undefined){
            let list = this.props.tables;
            let li = [];
            for (let i = 0; i < list.length; i++){
                li.push(<Table table = {list[i]}/>)
            }
            return (
                <ul className="sub-menu">{li}</ul>
            );
            
        }else{
            return (<ul className="sub-menu"></ul>);
        }

    }
  
}

function mapStateToProps(state) {
  const props = { workbook : state.workbook };
  return props;
}
function mapDispatchToProps(dispatch) {
  const actions = {};
  const actionMap = { actions: bindActionCreators(actions, dispatch) };
  //console.log(actionMap.actions.addTodo);
  return actionMap;
}
TableList = connect(mapStateToProps, mapDispatchToProps)(TableList)

export default TableList

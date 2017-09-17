import React from 'react'
import { connect } from 'react-redux'
//import { Fabric } from 'office-ui-fabric-react/lib/Fabric'
import Toggle from 'react-toggle'
import {  displaytips, displaypaint } from '../actions'
import { LoadWorkbook,ListenWorkbook } from '../actions/ExcelAction'
import { bindActionCreators } from 'redux';
import ReactCSSTransitionGroup from 'react-addons-css-transition-group'
import  Sheet  from './Sheet'
import "react-toggle/style.css" 
class SheetList extends React.Component{
    constructor(props){
        super(props);
        this.toogleValue = false;
        this.props.actions.ListenWorkbook();
        this.props.actions.LoadWorkbook();
    }
           
     
    render(){
        
        if (this.props.workbook.worksheets != undefined){
            let list = this.props.workbook.worksheets;
            let li = [];
            for (let i = 0; i < list.length; i++){
                li.push(<Sheet key={list[i].id} worksheet = {list[i]}/>)
            }
            return (
            <div>
                <div style = {{height :'40px', borderBottom : '1px solid #227447',borderTop : '5px solid #e6e6e6'}}>
                <div style = {{height :'30px',marginLeft: '30px', width : "20%", float : 'left'}}>

                   <label>
                    <Toggle
                        defaultChecked={this.toogleValue}
                        onChange={()=>{
                            
                            this.toogleValue= !this.toogleValue,
                            
                            this.props.actions.displaypaint({paint :this.toogleValue});
                        }} />
                    
                    </label>
                    
                </div>  
                    <div style = {{marginTop: '5px'}}>
                        <span >&nbsp;&nbsp;Selection Color</span>
                    </div>
                </div>
                <ul className="menu-content">
   
                    <ReactCSSTransitionGroup
                        transitionEnterTimeout={300}
                        transitionLeaveTimeout={100}
                        transitionName="example">
                        {li}
                    </ReactCSSTransitionGroup>
                </ul>
                
            </div>
            );
            
        }else{
            return (<ul className="menu-content"></ul>);
        }

    }
  
}

function mapStateToProps(state) {
  const props = { workbook : state.workbook };
  return props;
}
function mapDispatchToProps(dispatch) {
  const actions = {LoadWorkbook,ListenWorkbook,displaytips,displaypaint};
  const actionMap = { actions: bindActionCreators(actions, dispatch) };
  //console.log(actionMap.actions.addTodo);
  return actionMap;
}
SheetList = connect(mapStateToProps, mapDispatchToProps)(SheetList)

export default SheetList

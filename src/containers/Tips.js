import React from 'react'
import { connect } from 'react-redux'
import { displaytips } from '../actions'
import { bindActionCreators } from 'redux'
class Tips extends React.Component{
    constructor(props){
        super(props);
        //console.log(props);
        this.timeOutTips = undefined;
    }
    
    componentWillUpdate(nextProps, nextState){
        if (nextProps.display.active){
            if (this.timeOutTips) clearTimeout(this.timeOutTips);
            let props = this.props;
            this.timeOutTips = setTimeout(function(){
                props.actions.displaytips({text:nextProps.display.text,active:false})
            },2000);
        }
        
    }
    render(){
        //for (let i = 0; i < workbook)
        return (
            <div id = "tips" className={this.props.display.active?"active":""}>
                {this.props.display.text}
            </div>         
        );
    }
}
function mapStateToProps(state) {
  const props = { display : state.display };
  return props;
}
function mapDispatchToProps(dispatch) {
  const actions = {displaytips};
  const actionMap = { actions: bindActionCreators(actions, dispatch) };
  //console.log(actionMap.actions.addTodo);
  return actionMap;
}
Tips = connect(mapStateToProps, mapDispatchToProps)(Tips)

export default Tips

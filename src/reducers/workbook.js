const workbook = (state = {}, action) => {
  switch (action.type) {
    case 'HOSTUPDATE_SET_WORKBOOK':{
        
      return action.state;
    }
    case 'HOSTUPDATE_SET_TABLES_AND_CHARTS':{
        let newstate = {...state,worksheets : []};
        for (let i = 0; i < state.worksheets.length; i++){
            if (state.worksheets[i].id != action.state.id){
                newstate.worksheets.push(state.worksheets[i]);
            }else{
                newstate.worksheets.push(action.state);
            }
        }
        
        return newstate;
    }    
    case 'HOSTUPDATE_WORKSHEET_DELETED':
    {
        let newstate = {...state,worksheets : []};
        for (let i = 0; i < state.worksheets.length; i++){
            if (state.worksheets[i].name != action.id){
                newstate.worksheets.push(state.worksheets[i]);
            }
        }
        return newstate;
    }

    case 'HOSTUPDATE_WORKSHEET_ADDED':{
        let newstate = {...state,worksheets : []};
        for (let i = 0; i < state.worksheets.length; i++){
            newstate.worksheets.push(state.worksheets[i]);
        }
        newstate.worksheets.push({name:action.worksheet.name,id:action.worksheet.id,tables:[],charts:[]});
        return newstate;
    }
    case 'HOSTUPDATE_WORKSHEET_ACTIVATED':{
        let newstate = {...state, activesheet : action.worksheet};
        return newstate;
    }

    case 'HOSTUPDATE_WORKSHEET_DEACTIVATED':{
        return state;
    }

    case 'HOSTUPDATE_WORKSHEET_DATACHANGED':{
        let newstate = {...state,worksheets : []};
        for (let i = 0; i < state.worksheets.length; i++){
            if (state.worksheets[i].id != action.state.worksheet.id){
                newstate.worksheets.push(state.worksheets[i]);
            }else{
                let tmp = state.worksheets[i];
                tmp.datachangedstate = action.state.datachangedstate;
                newstate.worksheets.push(tmp);
            }
        }
        return newstate;
    }
    case 'HOSTUPDATE_WORKSHEET_SELECTIONCHANGED':{
        let newstate = {...state,worksheets : []};
        for (let i = 0; i < state.worksheets.length; i++){
            if (state.worksheets[i].id != action.state.worksheet.id){
                newstate.worksheets.push(state.worksheets[i]);
            }else{
                let tmp = {...state.worksheets[i]};
                tmp.selectionaddress = action.state.range;
                newstate.worksheets.push(tmp);
            }
        }
        return newstate;
    }
    case 'HOSTUPDATE_ACTIVATE_TABLE':{
        let newstate = {...state, activetableid : action.state.tableid};
        console.log(newstate);
        return newstate;
    }
    default:
      return state
  }
}

export default workbook;

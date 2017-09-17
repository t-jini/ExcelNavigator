import { store } from '..'
import { LoadWorkbook,actiondone,SetWorkbook,SetTablesAndCharts,SelectionchangedWorksheet,DatachangedWorksheet,DeleteWorksheet, AddWorksheet, ActivateWorksheet, DeactivateWorksheet } from '../actions/Excel'
import { ExcelAdapter } from '../services/Listener'

const config = (state = {nums : 0, actions : []}, action) => {
    let newstate = [];
    for (let i = 0; i < state.actions.length; i++) if (!state.actions[i].done) newstate.push(state.actions[i]);
    newstate.push({...action,id : state.nums});
    
  switch (action.type) {
    case 'ACTION_DONE' :{
        for (let i = 0; i < state.actions.length; i++)if (state.actions[i].id == action.id){
            state.actions[i].done = true;
            break;
        }
        return state;
    }

    case 'LISTEN_WORKBOOK':{
        ExcelAdapter.Adapter_ListenWorkbook(action,state);
        return {nums : state.nums + 1, actions : newstate};
    }           
    case 'LISTEN_WORKSHEET':
    {
        ExcelAdapter.Adapter_ListenWorksheet(action,state);
        return {nums : state.nums + 1, actions : newstate};
    }

    case 'ADD_TABLESANDCHARTS':{
        ExcelAdapter.Adapter_AddTablesANDChartes(action,state);
        return {nums : state.nums + 1, actions : newstate};
    }
    case 'DELETE_WORKSHEET':{
        ExcelAdapter.Adapter_DeleteWorksheet(action,state);
        return {nums : state.nums + 1, actions : newstate};
    }
    case 'ACTIVATE_WORKSHEET':{
        ExcelAdapter.Adapter_ActivateWorksheet(action,state);
        return {nums : state.nums + 1, actions : newstate};
    }
    case 'COLORED_RANGE':{
        ExcelAdapter.Adapter_ColorRange(action,state);
        return {nums : state.nums + 1, actions : newstate};
    }    

    case 'LOAD_WORKBOOK':{
        ExcelAdapter.Adapter_LoadWorkbook(action,state);
        return {nums : state.nums + 1, actions : newstate};
    }
    
    default:{
        return state;
    }
  }
}

export default config;

const display = (state = {}, action) => {
  switch (action.type) {
    case 'DISPLAY_TIPS':{
      return {text : action.state.text, active : action.state.active, paint : state.paint}
    }
    case 'DISPLAY_PAINT':{
      return {text : state.text, active : state.active, paint : action.state.paint}
    }
    default:
      return state
  }
}

export default display

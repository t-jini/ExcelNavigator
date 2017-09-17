import { combineReducers } from 'redux'
import workbook from './workbook'
import display from './display'
import config from './config'
const todoApp = combineReducers({
  workbook,
  display,
  config
})

export default todoApp

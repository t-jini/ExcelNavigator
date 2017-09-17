import React from 'react'
import { render } from 'react-dom'
import { createStore } from 'redux'
import { Provider } from 'react-redux'
import App from './components/App'
import reducer from './reducers'
import { ExcelAdapter } from './services/Listener'
export const store = createStore(reducer)
ExcelAdapter.setStore(store);
window.Office.initialize = () => {
  
  render(
  <Provider store={store}>
    <App />
  </Provider>,
  document.getElementById('root')
  )
};


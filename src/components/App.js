import React from 'react'
import SheetList from '../containers/SheetList'
import ExcelModule from './ExcelModule'
import Tips from '../containers/Tips'


class App extends React.Component{

  render(){
    return( 
      <div>
        
        <div className="nav-side-menu">
          <div className="brand"></div>
          <div className="menu-list">
            <SheetList />
            <ul id="menu-content" className="menu-content"></ul>
          </div>
          <Tips />
        </div>
        
    </div>);
  }
}


export default App

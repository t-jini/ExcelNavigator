


export const HostUpdate_SetWorkbook = (state) => ({
  type: 'HOSTUPDATE_SET_WORKBOOK',
  state
})

export const HostUpdate_SetTablesAndCharts = (state) => ({
  type: 'HOSTUPDATE_SET_TABLES_AND_CHARTS',
  state
})

export const HostUpdate_WorksheetDeleted = (id) => ({
  type: 'HOSTUPDATE_WORKSHEET_DELETED',
  id
})

export const HostUpdate_WorksheetAdded = (worksheet) => ({
  type: 'HOSTUPDATE_WORKSHEET_ADDED',
  worksheet
})

export const HostUpdate_WorksheetActivated = (worksheet) => ({
  type: 'HOSTUPDATE_WORKSHEET_ACTIVATED',
  worksheet
})

export const HostUpdate_WorksheetDeactivated = (worksheet) => ({
  type: 'HOSTUPDATE_WORKSHEET_DEACTIVATED',
  worksheet
})

export const HostUpdate_WorksheetSelectionchanged = (state) => ({
  type: 'HOSTUPDATE_WORKSHEET_SELECTIONCHANGED',
  state
})

export const HostUpdate_WorksheetDatachanged = (state) => ({
  type: 'HOSTUPDATE_WORKSHEET_DATACHANGED',
  state
})

export const HostUpdate_ActivateTable = (state) => ({
  type: 'HOSTUPDATE_ACTIVATE_TABLE',
  state
})


export const actiondone = (id) => ({
  type: 'ACTION_DONE',
  id
})
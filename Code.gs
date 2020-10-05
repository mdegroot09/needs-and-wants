function sheetNames(){
  let out = new Array()
  let sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  for (let i = 0; i < sheets.length; i++){
    out.push(sheets[i].getName())
  }
  out.splice(0,1)
  out.splice(-1,1)
  return out
}

function addToSprint(){
  let userEmail = Session.getActiveUser().getEmail()
  if (userEmail !== 'mike.degroot@homie.com' && userEmail !== 'james.wilson@homie.com'){
    return // quit if user is not Mike or James
  }
  
  let ss = SpreadsheetApp.getActive().getActiveSheet()
  let vertical = ss.getSheetName()
  let task = ss.getRange('G2').getValue()
  let oldRow = ss.getRange('J2').getValue()
  if (oldRow == 0){return} // quit if oldRow is not found
  let vals = ss.getRange(oldRow + ':' + oldRow).getValues()
  ss.deleteRow(oldRow)
  let newRow = findNewRow()
  ss.getRange(newRow + ':' + newRow).setValues(vals)
  ss.getRange('G2').setValue('')
  return addToMainSheet(task, vertical)
}

function findNewRow(){
  let ss = SpreadsheetApp.getActive().getActiveSheet()
  let newRow = 4
  while (true){
    let val = ss.getRange('A' + newRow).getValue()
    if (val == 'Future Sprints'){
      ss.insertRowBefore(4) // add row if none is available
      return 4 
    }
    else if (val){ // skip iteration if row is not blank
      newRow += 1
      continue
    } 
    else {
      return newRow
    }
  }
}

function addToMainSheet(task, vertical){
  let ss = SpreadsheetApp.getActive().getSheetByName('Current Sprint')
  ss.insertRowBefore(3)
  ss.getRange('A3:B3').merge()
  ss.getRange('A3').setValue(task)
  ss.getRange('C3:D3').merge()
  ss.getRange('C3').setValue(vertical)
  ss.getRange('E3').setValue('=IF($A6 <> "", VLOOKUP($A3,INDIRECT($C3&"!A4:F"),2,false), "")')
  ss.getRange('F3').setValue('=IF($A6 <> "", VLOOKUP($A3,INDIRECT($C3&"!A4:F"),3,false), "")')
  ss.getRange('G3').setValue('=IF($A6 <> "", VLOOKUP($A3,INDIRECT($C3&"!A4:F"),4,false), "")')
  ss.getRange('H3').setValue('=IF($A6 <> "", VLOOKUP($A3,INDIRECT($C3&"!A4:F"),5,false), "")')
  ss.getRange('I3').setValue('=IF($A6 <> "", VLOOKUP($A3,INDIRECT($C3&"!A4:F"),6,false), "")')
}

function archiveCompleted(){
  let currentSprint = SpreadsheetApp.getActive().getSheetByName('Current Sprint')
  let archive = SpreadsheetApp.getActive().getSheetByName('Archive')
  let rows = []
  currentSprint.getRange('A3:M').getValues().forEach(function(a, i){
    if (a[10] == 'Complete'){
      let vals = currentSprint.getRange('A' + (i + 3)).getValues()
      archive.insertRowBefore(3)
      archive.getRange('A3:B3').merge()
      archive.getRange('A3').setValue(a[0])
      archive.getRange('C3:D3').merge()
      archive.getRange('C3').setValue(a[2])
      archive.getRange('E3').setValue(a[4])
      archive.getRange('F3').setValue(a[5])
      archive.getRange('G3').setValue(a[6])
      archive.getRange('H3').setValue(a[7])
      archive.getRange('I3').setValue(a[8])
      archive.getRange('J3').setValue(a[9])
      archive.getRange('K3').setValue(a[10])
      archive.getRange('L3').setValue(a[11])
    }
  })
  
  return deleteComplete()
}

function deleteComplete(){
  let ss = SpreadsheetApp.getActive().getSheetByName('Current Sprint')
  let row = 3
  while (ss.getRange('A' + row).getValue()){
    if (ss.getRange('K' + row).getValue() == 'Complete'){
      ss.deleteRow(row)
      continue
    }
    row += 1
  }
}
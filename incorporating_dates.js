/**
 * Copies source range and pastes at first empty row on target sheet
 */
function CopyIt(){
  var source_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var target_spreadsheet = SpreadsheetApp.openById("My_Target_Sheet_Key");
  

  // Get source and target sheets - can be the same or different
  var sourcesheet = source_spreadsheet.getSheetByName("Form Responses");
  var targetsheet = target_spreadsheet.getSheetByName("Work_Orders");

  //Get row of last form submission
  var source_last_row = sourcesheet.getLastRow();

  // Check for answer to Flyer Created? If Yes, end now. If not No, i.e. Yes, continue.
  var check = sourcesheet.getRange("T"+(source_last_row)).getValue();
  Logger.log(check);
  if (check != 'Yes'){  
  
  // Get the source ranges
  // TimeStamp
    var source_range1 = sourcesheet.getRange("A"+(source_last_row));
  
  //Name
    var source_range2 = sourcesheet.getRange("B"+(source_last_row));
 
    
  //Phone Number
    var source_range3 = sourcesheet.getRange("C"+(source_last_row));
  
    
  //Class Name/Title
    var source_range4 = sourcesheet.getRange("E"+(source_last_row));
  
    
  //First Session
    var source_range5 = sourcesheet.getRange("D"+(source_last_row));
    
  //Dates and Times
    
    var source_range6 = sourcesheet.getRange("H"+(source_last_row));
    var source_range7 = sourcesheet.getRange("I"+(source_last_row));
    var source_range8 = sourcesheet.getRange("J"+(source_last_row));
    var source_range9 = sourcesheet.getRange("K"+(source_last_row));
    var source_range10 = sourcesheet.getRange("L"+(source_last_row));
    var source_range11 = sourcesheet.getRange("M"+(source_last_row));
    var source_range12 = sourcesheet.getRange("N"+(source_last_row));
    var source_range13 = sourcesheet.getRange("O"+(source_last_row));
    var source_range14 = sourcesheet.getRange("P"+(source_last_row));
    var source_range15 = sourcesheet.getRange("Q"+(source_last_row));
    
    var source_range1_values = source_range1.getValues();    
    var source_range2_values = source_range2.getValues();    
    var source_range3_values = source_range3.getValues();    
    var source_range4_values = source_range4.getValues();       
    var source_range5_values = source_range5.getValues();
    var source_range6_values = source_range6.getValues();
    var source_range7_values = source_range7.getValues();
    var source_range8_values = source_range8.getValues();
    var source_range9_values = source_range9.getValues();
    var source_range10_values = source_range10.getValues();
    var source_range11_values = source_range11.getValues();
    var source_range12_values = source_range12.getValues();
    var source_range13_values = source_range13.getValues();
    var source_range14_values = source_range14.getValues();
    var source_range15_values = source_range15.getValues();
    
  // Get the last row on the target sheet
  var last_row = targetsheet.getLastRow();
  
  // Set the target ranges on target sheet
  //TimeStamp
  var target1 = targetsheet.getRange("A"+(last_row+1));
    
  //Name
  var target2 = targetsheet.getRange("H"+(last_row+1));
    
  //Phone
  var target3 = targetsheet.getRange("I"+(last_row+1));
    
  //Your Project Name
  var target4 = targetsheet.getRange("K"+(last_row+1));
    
  //Final Due Date
  var target5 = targetsheet.getRange("M"+(last_row+1));
    
  //The next two are defined in script
  var target6 = targetsheet.getRange("O"+(last_row+1));   
  var target7 = targetsheet.getRange("Q"+(last_row+1));
  
    // Just using the text to put this data, so there is no source range above.   
  target6.setValue('Flyer');
  target7.setValue('This job was imported from a Org Learning\'s Class Reservation Form and followup will be required for clarification'); 
    
    //Date and Time
    var target8 = targetsheet.getRange("U"+ " "+ (last_row+1));
    var target9 = targetsheet.getRange("U"+ " "+ (last_row+1));
    var target10 = targetsheet.getRange("U"+ " "+ (last_row+1));
    var target11 = targetsheet.getRange("U"+ " "+ (last_row+1));
    var target12 = targetsheet.getRange("U"+ " "+ (last_row+1));
    var target13 = targetsheet.getRange("U"+ " "+ (last_row+1));
    var target14 = targetsheet.getRange("U"+ " "+ (last_row+1));
    var target15 = targetsheet.getRange("U"+ " "+ (last_row+1));
    var target16 = targetsheet.getRange("U"+ " "+ (last_row+1));
    var target17 = targetsheet.getRange("U"+ " "+ (last_row+1));
	
	
  // Put the data from the source sheet into the target sheet, after adding a new row
    targetsheet.insertRowAfter(last_row);
    target1.setValues(source_range1_values);
    target2.setValues(source_range2_values);
    target3.setValues(source_range3_values);
    target4.setValues(source_range4_values);
    target5.setValues(source_range5_values);
    target8.setValues(source_range6_values);
    target9.setValues(source_range7_values);
    target10.setValues(source_range8_values);
    target11.setValues(source_range9_values);
    target12.setValues(source_range10_values);
    target13.setValues(source_range11_values);
    target14.setValues(source_range12_values);
    target15.setValues(source_range13_values);
    target16.setValues(source_range14_values);
    target17.setValues(source_range15_values);
 
  }
}  

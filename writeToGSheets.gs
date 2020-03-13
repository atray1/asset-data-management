
/*** This function is called whenever a form is submitted
*/
function getLastRow(s) {

   var s =   s || SpreadsheetApp.getSheetByName('Form Responses 1')
   var  v = s.getRange('A:A').getValues(),
        l = v.length,
        r;
    while (l > 0) {
        if (v[l] && v[l][0].toString().length > 0) {
            r = (l + 1);
            break;
        } else {
            l--;
        }
    }
    Logger.log(r)
    return r;
}

function onSubmit(e){
  var response = e.values;
  
  //first determine the action
  var action = response[form_dict.action]; 
  if(action == undefined || action == '') { 
    action = response[form_dict.recyc_only]; 
  }
  
  //then determine the asset
  var asset;
  var store_bool = false;
  switch(action){
    case 'Reuse/Sell/Donate (SAVE program - Surplus Asset Value rEalization)':
      asset = response[form_dict.SAVE_asset];
      break;
    case 'Recycle (broken items)':
    case 'Recycle (broken items, junk)':
      asset = response[form_dict.recyc_asset];
      break;
    case 'Store (for later use)':
      asset = response[form_dict.store_asset];
      store_bool = true;
      break;
    default:
      Logger.log("ERROR ACTION: "+action);
  }
  
  //grabbind SD number, for actual sheet
  var lr = getLastRow(form_s);
  var sd_numm = form_s.getRange(lr, form_dict.sdnum+1).getValue();
  var importSheet = master_s
  importSheet.getRange(importSheet.getLastRow() + 1 , 18).setValue(sd_numm); 
  //grabbing line number, for actual sheet
  var line_numm = form_s.getRange(lr, form_dict.linenum+1).getValue();
  importSheet.getRange(importSheet.getLastRow() +0 , 23).setValue(line_numm);
  
  // sd num for copy
  var importSheets = master_copy
  importSheets.getRange(importSheets.getLastRow() + 1 , 18).setValue(sd_numm);
  // line num for copy
  importSheets.getRange(importSheets.getLastRow()+0 , 23).setValue(line_numm);
  
  //write to dummy sheet Master
  var master_data = [response[form_dict.Timestamp], asset, action, response[form_dict.store_len]];
  var model = '';
  var serial = '';
  var ein = '';
  var asset_tag = '';
  if(asset == 'Laboratory Equipment'){
    var lab_info1 = response.slice(form_dict.lab_desc, form_dict.lab_manu+1);
    var lab_info2 = response.slice(form_dict.lab_building, form_dict.lab_email+1);
    var lab_info3 = response.slice(form_dict.lab_cc, form_dict.lab_hazmat+1);
    master_data = master_data.concat(lab_info1, lab_info2, lab_info3, [response[form_dict.lab_data]]);
    model = response[form_dict.lab_model];
    serial = response[form_dict.lab_serial];
    ein = response[form_dict.lab_ein];
    asset_tag = response[form_dict.asset_tag];
  } else if (asset == 'Miscellaneous (other)'|| asset == 'Monitor, Keyboard, Mouse'){
    var misc_info1 = response.slice(form_dict.misc_desc, form_dict.misc_manu+1);
    var misc_info2 = response.slice(form_dict.misc_building, form_dict.misc_email+1);
    var misc_info3 = response.slice(form_dict.misc_cc, form_dict.misc_super+1);
    master_data = master_data.concat(misc_info1, misc_info2, misc_info3, [response[form_dict.misc_hazmat]]);
    serial = response[form_dict.misc_serial];
  }
  var row_num = writeToSD(master_data, master_s);
  
  // write to master copy
  var row_numm = writeToSD(master_data, master_copy);
  
  //if cpu is involved. stop here
  if(asset == 'Recycle/Dispose Laptop and Desktop Computers ONLY - EXCLUDING Servers'){ return; }
  
  //write to respective sheet
  var answer;
  var location = master_data[dict.building]+' '+master_data[dict.room];
    
  //determine the action
  if(action == 'Reuse/Sell/Donate (SAVE program - Surplus Asset Value rEalization)'){
    //For SAVE
    answer = [line_numm, master_data[dict.manu], model, master_data[dict.desc], serial, ein, asset_tag, master_data[dict.cc], response[form_dict.condition], "", master_data[dict.super], "", "", "",
              sd_numm, "", "", response[form_dict.Timestamp], master_data[dict.special]];
    
    //check if it's lab equipment
    if(asset == 'Laboratory Equipment'){
      //check if it has asset tag
      if(asset_tag!='Non-Capital'){
        writeToSD(answer, cap_s);//capital
        writeToSD(answer, cap_ss);// copy capital
      } else {
        writeToSD(answer, ncap_s);//non-capital
        writeToSD(answer, ncap_ss);//copy non-capital
      }
    } else { //goes into miscellaneous
      writeToSD(answer, misc_s);
      writeToSD(answer, misc_ss);
    }
    
    //For Recycling
  } else if (action == 'Recycle (broken items)'|| action == 'Recycle (broken items, junk)') {
    answer = [line_numm, master_data[dict.manu], model, master_data[dict.desc], serial, ein, asset_tag, master_data[dict.cc], location, sd_numm, response[form_dict.Timestamp], response[form_dict.condition]];
    writeToSD(answer, recycle_s);
    
    answer = [line_numm, master_data[dict.manu], model, master_data[dict.desc], serial, ein, asset_tag, master_data[dict.cc], location, sd_numm, response[form_dict.Timestamp], response[form_dict.condition]];
    writeToSD(answer, recycle_ss);
    
    //For Storage
  } else if (action == 'Store (for later use)'){
    answer = [line_numm, sd_numm, "", master_data[dict.manu], model, master_data[dict.desc], serial, ein, asset_tag, master_data[dict.cc], response[form_dict.condition], master_data[dict.super], master_data[dict.name]];
    writeToSD(answer, store_s);
    
    //this writes to storage copy
    answer = [line_numm, sd_numm, "", master_data[dict.manu], model, master_data[dict.desc], serial, ein, asset_tag, master_data[dict.cc], response[form_dict.condition], master_data[dict.super], master_data[dict.name]];
    writeToSD(answer, store_sss);
  }
  
  //bounce email
  var storage_len = (store_bool)? master_data[dict.store_len]: "Not Storage";  
  var fields = [sd_numm, action, master_data[dict.manu], ein, model, asset_tag, master_data[dict.desc], response[form_dict.condition], serial, response[form_dict.lab_data], master_data[dict.name], master_data[dict.Timestamp],
               master_data[dict.email], master_data[dict.cc], master_data[dict.super], location, master_data[dict.special]];
  Logger.log('got here');
  bounceEmail(fields, master_data[dict.hazmat], sd_numm);
  
}

/*** 
*/
function writeToSD(info, sheet){
  //find next available row
  var first_col = sheet.getRange(1, 2, sheet.getLastRow(), 1).getValues(); //some sheet has empty 1st column
  var first_col_T = first_col.map(function(r){return r[0]});
  var next_row = first_col_T.filter(String).length + 1;
  
  //write to the sheet
  sheet.getRange(next_row, 1, 1, info.length).setValues([info]);
  SpreadsheetApp.flush();
  
  return next_row;
}



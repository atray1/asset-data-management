/*** All the Global Variables ***/
//ID of the gSheet linked to the form
var id = '1kRAXmnjh2eG8ESem8_Mfnfn-WqLhRQBNkUCxtEGYl7k';
var spread_sheet = SpreadsheetApp.openById(id);
var form_s = spread_sheet.getSheetByName('Form Responses 1');
var master_s = spread_sheet.getSheetByName('Master Sheet - All');
var cap_s = spread_sheet.getSheetByName('Available Capital Assets');
var ncap_s = spread_sheet.getSheetByName('Available Non-Capital Assets');
var store_s = spread_sheet.getSheetByName('Storage'); //If I would like to write this storage sheet instead of the actual storage sheet
var recycle_s = spread_sheet.getSheetByName('Recycle');
var misc_s = spread_sheet.getSheetByName('SAVE-Miscellaneous');

//import to copy dummy sheets
var store_sss = SpreadsheetApp.openById('1u9DKDummsAq0JNtWPTj92no_pIDpa6WJEQ6PpaRZ0Jk').getSheetByName('Storage');
var master_copy = SpreadsheetApp.openById('1u9DKDummsAq0JNtWPTj92no_pIDpa6WJEQ6PpaRZ0Jk').getSheetByName('Master Sheet - All');
var cap_ss = SpreadsheetApp.openById('1u9DKDummsAq0JNtWPTj92no_pIDpa6WJEQ6PpaRZ0Jk').getSheetByName('Available Capital Assets');
var ncap_ss = SpreadsheetApp.openById('1u9DKDummsAq0JNtWPTj92no_pIDpa6WJEQ6PpaRZ0Jk').getSheetByName('Available Non-Capital Assets');
var misc_ss = SpreadsheetApp.openById('1u9DKDummsAq0JNtWPTj92no_pIDpa6WJEQ6PpaRZ0Jk').getSheetByName('SAVE-Miscellaneous');
var recycle_ss = SpreadsheetApp.openById('1u9DKDummsAq0JNtWPTj92no_pIDpa6WJEQ6PpaRZ0Jk').getSheetByName('Recycle');

/*** Dictionaries ***/
//form response dictionary 
var form_dict = {Timestamp:0, condition:1, recyc_only:2, action:3, SAVE_asset:4, store_asset:5, store_len:6, recyc_asset:7, misc_desc:8, misc_manu:9, misc_serial:10, misc_building:11,    misc_room:12,
                 misc_gmp:13, misc_special:14, misc_name:15, misc_phone:16, misc_email:17, misc_cc:18, misc_super:19, misc_second:20, misc_hazmat:21, asset_tag:22, lab_desc:23, lab_manu:24, lab_model:25,
                 lab_serial:26, lab_ein:27, lab_building:28, lab_room:29, lab_gmp:30, lab_special:31, lab_name:32, lab_phone:33, lab_email:34, lab_second:35, lab_cc:36, lab_super:37, lab_hazmat:38, lab_data:39, sdnum:40, linenum:41};

//Master Sheet dictionary
var dict = {Timestamp:0, asset:1, action:2, storage_len:3, desc:4, manu:5, building:6, room:7, gmp:8, special:9, name:10, phone:11, email:12, cc:13, super:14, hazmat:15, data:16, sd:17, line:22};

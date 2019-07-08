require('./activex');

var path = require('path');
var data_path = path.join(__dirname, '../data/');
var filename = "persons.dbf";
var constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + data_path + ";Extended Properties=\"DBASE IV;\"";

console.log("==> Preapre directory and delete DBF file on exists");
var fso = new ActiveXObject("Scripting.FileSystemObject");
if (!fso.FolderExists(data_path)) fso.CreateFolder(data_path);
if (fso.FileExists(data_path + filename)) fso.DeleteFile(data_path + filename);

console.log("==> Open connection");


//----------------------------------------------


var hwp = new ActiveXObject('HWPFrame.HwpObject.1');
hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet);	// 텍스트 입력
hwp.HParameterSet.HInsertText.Text = '테스트입니다.';
hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet);
hwp.HAction.Run('BreakPara');	// 엔터 입력
hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet);	// 텍스트 입력
hwp.HParameterSet.HInsertText.Text = '테스트입니다222';
hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet);
hwp.HAction.Run('SelectAll');	// 모두선택
hwp.HAction.Run('CharShapeBold');	// 진하게
hwp.HAction.Run('Cancel');	// 블록해제

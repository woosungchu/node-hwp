require('winax');

var hwp = new ActiveXObject('HWPFrame.HwpObject.1');
var hAct = hwp.HAction;
var filename = (function(){return Date.now()})();

hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule");
hwp.Open("C:/aProjects/node-hwp/resource/template1.hwp","HWP","template:true"); // 접근제한
hwp.SaveAs("C:/Users/admin/Downloads/test"+filename+".hwp","HWP","download:true")


//hAct.Run("Paste"); // 클립보드 내용 붙이기
/*
hAct.Run("MoveUp");
hAct.Run("MoveUp");
hAct.Run("MoveUp");
hAct.Run("MoveUp");
hAct.Run("MoveUp");
hAct.Run("MoveUp");
hAct.Run("MoveUp");
hAct.Run("MoveUp");
hAct.Run("MoveUp");
hAct.Run("MoveUp");
hAct.Run("MoveUp");
hAct.Run("MoveUp");
hAct.Run("MoveUp");
hAct.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
hwp.HParameterSet.HInsertText.Text = " 2019학년도 1학기                   고등학교  1학년  1반  1번  이름 : 한 상 욱";
hAct.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
hAct.Run("MoveLeft");
hAct.Run("MoveLeft");
hAct.Run("MoveLeft");
hAct.Run("MoveLeft");
hAct.Run("MoveLeft");
hAct.Run("MoveLeft");
hAct.Run("MoveLeft");
hAct.Run("MoveLeft");
hAct.Run("MoveLeft");
hAct.Run("MoveLeft");
hAct.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
hwp.HParameterSet.HInsertText.Text = "가나다 수준에 따른 평가";
hAct.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
*/

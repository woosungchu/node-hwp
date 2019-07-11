require('winax');
var fn = require('./fn.js');

var hwp = new ActiveXObject('HWPFrame.HwpObject.1');
var hAct = hwp.HAction;
var filename = (function(){return Date.now()})();

hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule");
hwp.Open("C:/aProjects/node-hwp/resource/template1.hwp","HWP","template:true"); // 접근제한


// 내용 편집 시작

hAct.Run("MoveDocBegin");
fn.write(hwp,"진로와직업");
hAct.Run("MoveDown");
hAct.Run("MoveDown");
fn.write(hwp," 2019학년도 1학기                   고등학교  1학년  1반  1번  이름 : 한 상 욱");

hAct.Run("MoveDocEnd");
hAct.Run("MoveUp");
hAct.Run("MoveUp");
hAct.Run("MoveRight");
hAct.Run("MoveRight");
fn.write(hwp,"가나다 수준에 따른 평가");

// 내용 편집 끝


hwp.SaveAs("C:/Users/admin/Downloads/test"+filename+".hwp","HWP","download:true");

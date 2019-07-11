require('winax');
var fn = require('./fn.js');

var hwp = new ActiveXObject('HWPFrame.HwpObject.1');
var hAct = hwp.HAction;
var filename = Date.now();
var students = [{name:'기창빈',grade:'가'},{name:'한상욱',grade:'다'}];

hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule");

for (var i = 0; i < students.length; i++) {
  console.log(students[i].name,students[i].grade)
  filename = Date.now() + '_' + students[i].name + '_'+ students[i].grade;

  hwp.Open("C:/aProjects/node-hwp/resource/template1.hwp","HWP","template:true");

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
  fn.write(hwp,students[i].name + "은" + students[i].grade + '수준 입니다.');

  // 내용 편집 끝

  hwp.SaveAs("C:/Users/admin/Downloads/test"+filename+".hwp","HWP","download:true");
}

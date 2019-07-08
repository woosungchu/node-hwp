require('./activex');

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

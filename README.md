#### nodeHwp

https://github.com/durs/node-activex
https://joyfui.wo.tc/1124

https://www.hancom.com/board/devdataView.do?board_seq=47&artcl_seq=4084&pageInfo.page=1&search_text=

```
npm install -g windows-build-tools
npm install -g winax
npm install

node-gyp configure
node-gyp build
```


```
hwp.Open("C:/aProjects/node-hwp/test.hwp","HWP",null);  // 파일 열기

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
```

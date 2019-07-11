exports.write = function(hwp, text){
  hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
  hwp.HParameterSet.HInsertText.Text = text;
  hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
  return true;
}

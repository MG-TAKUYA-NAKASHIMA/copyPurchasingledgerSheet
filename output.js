//01→02への出力
function outputTriger() {
	let outputData = getdata();
	exportData(outputData);
}

//データの取得
function getdata() {
	const outputDataInNumbering = getOutputDataInNumbering();//01の「【出力シート】請求書(明細別)」のシートを特定
	let valueOfOutputData = outputDataInNumbering.getDataRange().getValues();//01の「【出力シート】請求書(明細別)」のシートデータを全件取得
	valueOfOutputData.splice(0,2);//01の「【出力シート】請求書(明細別)」のシートデータの見出し2行を削除
	return valueOfOutputData
}

//データの出力
function exportData(outputData) {
	const classifySheet = getClassifySheet();//02の「請求書(明細別)」のシートを特定
	let lastRow = classifySheet.getLastRow(),//02の「請求書(明細別)」のシートデータの最終行を取得
	lastCol = classifySheet.getLastColumn();//02の「請求書(明細別)」のシートデータの最終列を取得
	classifySheet.getRange(3, 1, lastRow, lastCol).clearContent();//02の「請求書(明細別)」のシートデータのみを削除
	classifySheet.getRange(3, 1, outputData.length, outputData[0].length).setValues(outputData);//02の「請求書(明細別)」に出力
}
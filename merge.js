// 請求書（明細別）データ入力で使用するスクリプト
//仕入先台帳への記入部分は未作成

//「【入力シート】請求書(明細別)」のデータと
//「カスタムID未採番者リスト」のデータを合わせて
//すべての人に仕入先code(カスタムID)を採番したデータを作成する
function reflectTriger() {
const mergeData = generateMergeData();
exportMergeData(mergeData);
}

function generateMergeData() {
	let valueOfInputData        = getInputDataInNumbering().getDataRange().getValues();
	let formatedData            = getUnnumberingPersonSheetInNumbering().getDataRange().getValues();

	for(let i = 2; valueOfInputData.length > i; i++) {
		if(valueOfInputData[i][13] === '') {
			for(let j = 2; formatedData.length > j; j++) {
				if(valueOfInputData[i][12] === formatedData[j][0]) {
					valueOfInputData[i][13] = formatedData[j][1];
				}
			}
		}
	}
	valueOfInputData.shift();
	valueOfInputData.shift();
	return valueOfInputData;
}

//「【出力シート】請求書(明細別)」シートへ貼付
function exportMergeData(mergeData) {
	const outputDataInNumbering = getOutputDataInNumbering();
	outputDataInNumbering.getRange(3,1,mergeData.length,mergeData[0].length).setValues(mergeData);
}
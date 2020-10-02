// 請求書（明細別）データ入力で使用するスクリプト
//仕入先台帳への記入部分は未作成

//「【入力シート】請求書(明細別)」のデータと
//「カスタムID未採番者リスト」のデータを合わせて
//すべての人に仕入先code(カスタムID)を採番したデータを作成する
function reflectTriger() {
	const outputDataInNumbering = getOutputDataInNumbering();
	let lastRow = outputDataInNumbering.getLastRow();
	outputDataInNumbering.getRange(3,1,lastRow,44).clearContent();

	let unNumberingList = generateNumberedList();
	const mergeData = generateMergeData(unNumberingList);
	mergeData.shift();
	exportMergeData(mergeData);
}

//採番された人のリストを「カスタムID未採番者リスト」と「採番済みリスト」から取得する
//変数のつけ方があまい
//2つの配列の変数名はそれぞれ別のものにするべきだしマージした配列は新しく宣言しなくてはいけない
function generateNumberedList() {
	let unNumberingList = getUnnumberingPersonSheetInNumbering().getDataRange().getValues(),
			numberedList    = getNumberedSheet().getDataRange().getValues();

	//見出しを削除する
	unNumberingList.shift();
	unNumberingList.shift();
	numberedList.shift();

	exportFormateData(unNumberingList);
	exportNumberedSheet(unNumberingList);


for(let i = 0; numberedList.length > i; i++) {
	unNumberingList.push(numberedList[i]);
}
return unNumberingList;

}

function generateMergeData(unNumberingList) {
	let valueOfInputData = getInputDataInNumbering().getDataRange().getValues();

	for (let i = 2; valueOfInputData.length > i; i++) {
		if (valueOfInputData[i][13] == '') {
			for (let j = 0; unNumberingList.length > j; j++) {
				if (valueOfInputData[i][12] == unNumberingList[j][1]) {
					valueOfInputData[i][13] = unNumberingList[j][0];
				} 
			}
		}
	}

	valueOfInputData.shift().shift();
	return valueOfInputData;
}

//「【出力シート】請求書(明細別)」シートへ貼付
function exportMergeData(mergeData) {
	const outputDataInNumbering = getOutputDataInNumbering();
	outputDataInNumbering.getRange(3, 1, mergeData.length, mergeData[0].length).setValues(mergeData);
}

//「採番済みリスト」シートへの書き込み
function exportNumberedSheet() {
	
		const numberedSheet = getNumberedSheet(),
		lastRow = numberedSheet.getLastRow(),
		unNumberingListSheet = getUnnumberingPersonSheetInNumbering();
		let unNumberingList = unNumberingListSheet.getDataRange().getValues();
		unNumberingList.shift();
		unNumberingList.shift();
		if(unNumberingList.length >= 1) {
		numberedSheet.getRange(lastRow + 1,1,unNumberingList.length,3).setValues(unNumberingList);
		unNumberingListSheet.getRange(3,1,unNumberingList.length,3).clearContent();

	}


}


//仕入先台帳へ反映
function exportFormateData(unNumberingList) {
	const supplierLedgerSheet = getSupplierLedgerSheet();
	let valueOfSupplierLedgerSheet = supplierLedgerSheet.getDataRange().getValues();
	let tmp;

	unNumberingList.forEach(arr => {
		arr.shift();
		tmp = arr[0];
		arr.push(tmp);
		arr.shift();
	});

	valueOfSupplierLedgerSheet.some((arr, i, self) => {
		if(unNumberingList.length > 0 && self[i][2] === unNumberingList[0][0] - 1 ) {
			supplierLedgerSheet.insertRows(i+2,unNumberingList.length);
			supplierLedgerSheet.getRange(i+2,3,unNumberingList.length,2).setValues(unNumberingList);
			return true;
		}
	})
}


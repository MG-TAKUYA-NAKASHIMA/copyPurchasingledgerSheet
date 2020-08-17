//「採番用シート」で使用するスクリプト

//「【入力シート】請求書(明細別)」シートの「カスタムID未割当者抽出」ボタンを押すと実行する
function triger() {
	let unNumberingData = searchUnnumberedPerson();
	unNumberingData     = deleateDuplicate(unNumberingData);
	let formatedData    = formatData(unNumberingData);
	exportUunumberingData(formatedData);
}


//「【入力シート】請求書(明細別)」に貼り付けられたデータのからカスタムID未採番者を特定する
function searchUnnumberedPerson() {
	const valueOfInputData = getInputDataInNumbering().getDataRange().getValues();
	let unNumberingData    = [];

	for (let i = 2; valueOfInputData.length > i; i++) {
		if (valueOfInputData[i][13] === '') {
			unNumberingData.push(valueOfInputData[i]);
		}

	}

	return unNumberingData;
}

//重複を削除する
function deleateDuplicate(unNumberingData) {
	unNumberingData = unNumberingData.filter(function(e, index){
		return !unNumberingData.some(function(e2, index2){
			return index > index2 && e[12] == e2[12];
		});

	});
	return unNumberingData;
}

//「pasture表示名(請求元)」だけ抜き出す
function formatData(unNumberingData) {
	let formatedData = [];
	for (let i = 0; unNumberingData.length > i; i++) {
		formatedData.push([unNumberingData[i][12]]);
	}

	return formatedData;
}

//「カスタムID未採番者リスト」シートに貼付を行う
function exportUunumberingData(formatedData) {
	const unnumberingPersonSheetInNumbering = getUnnumberingPersonSheetInNumbering();
	unnumberingPersonSheetInNumbering.getRange(3,1,formatedData.length,1).setValues(formatedData);
}
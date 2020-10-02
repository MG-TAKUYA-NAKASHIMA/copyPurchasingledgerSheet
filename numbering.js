//「請求書（明細別）データ入力」で使用するスクリプト
// 仕入先台帳から最新の仕入先codeを引っ張ってきている部分のスクリプトは未作成


//「【入力シート】請求書(明細別)」シートの「カスタムID未割当者抽出」ボタンを押すと実行する
function exportTriger() {
	let unNumberingData = searchUnnumberedPerson();
	unNumberingData     = deleateDuplicate(unNumberingData);
	let latestCustomId  = findLatestCustomId();
	let formatedData    = formatData(unNumberingData,latestCustomId);
	formatedData        = compareNumbered(formatedData)
	exportUunumberingData(formatedData);
}

function deleteTrigerInInput() {
	const inputDataInNumbering = getInputDataInNumbering();
	let lastRow = inputDataInNumbering.getLastRow();
	inputDataInNumbering.getRange(3,1,lastRow,44).clear();
}


//「【入力シート】請求書(明細別)」に貼り付けられたデータの中からカスタムID未採番者を特定する
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

//「pasture表示名(請求元)」だけ抜き出し、加工する
function formatData(unNumberingData,latestCustomId) {
	let tmp          = [];
	let formatedData = [];
	
	for (let i = 0; unNumberingData.length > i; i++) {
		tmp.push(unNumberingData[i][11]);
		tmp.push(unNumberingData[i][12]);
		tmp.push(latestCustomId);
		formatedData.push(tmp);
		tmp = [];
		latestCustomId++;
	}
	return formatedData;
}

//「採番済みリスト」と照合し、採番済みの人を除外する
function compareNumbered(formatedData) {
	const numberedList = getNumberedSheet().getDataRange().getValues();
	let deleateRows = [];

	for(let i = 0; formatedData.length > i; i++) {
		for(let c = 1; numberedList.length > c; c++) {
			if(formatedData[i][0] === numberedList[c][0]) {
				deleateRows.push(i);
			}
		}
	}

	for (let j = 0; deleateRows.length > j; j++) {
		formatedData.splice(deleateRows[j] - j, 1)
	}

	return formatedData;
}

//「カスタムID未採番者リスト」シートに貼付を行う
function exportUunumberingData(formatedData) {
	const unnumberingPersonSheetInNumbering = getUnnumberingPersonSheetInNumbering();
	if(formatedData.length > 0){
		unnumberingPersonSheetInNumbering.getRange(3,1,formatedData.length,3).setValues(formatedData);
	}
}

//「仕入先台帳」から最新の空き番号を取得する
function findLatestCustomId() {
	const supplierLedgerSheet        = getSupplierLedgerSheet();
	const valueOfsupplierLedgerSheet = supplierLedgerSheet.getDataRange().getValues();
	let latestCustomId;

	valueOfsupplierLedgerSheet.some((arr, i, self)=> {
		if(typeof self[i][1] == 'number' && self[i][1] === 0){
			latestCustomId = self[i - 1][2] + 1;
			return true;
		}
	});
	return latestCustomId;
}

 
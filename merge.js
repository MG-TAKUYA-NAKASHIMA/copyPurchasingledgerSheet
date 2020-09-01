// 請求書（明細別）データ入力で使用するスクリプト
//仕入先台帳への記入部分は未作成

//「【入力シート】請求書(明細別)」のデータと
//「カスタムID未採番者リスト」のデータを合わせて
//すべての人に仕入先code(カスタムID)を採番したデータを作成する
function reflectTriger() {
	let formatedData = getFormatedData();
	formatedData.shift();
	const mergeData = generateMergeData(formatedData);
	mergeData.shift();
	exportMergeData(mergeData);
	exportFormateData(formatedData);
}

function getFormatedData() {
	let formatedData = getUnnumberingPersonSheetInNumbering().getDataRange().getValues();
	formatedData.shift().shift();
	return formatedData;
}

function generateMergeData(formatedData) {
	let valueOfInputData = getInputDataInNumbering().getDataRange().getValues();

	for (let i = 2; valueOfInputData.length > i; i++) {
		if (valueOfInputData[i][13] === '') {
			for (let j = 0; formatedData.length > j; j++) {
				if (valueOfInputData[i][12] === formatedData[j][1]) {
					valueOfInputData[i][13] = formatedData[j][2];
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

function exportFormateData(formatedData) {
	const supplierLedgerSheet = getSupplierLedgerSheet();
	let valueOfSupplierLedgerSheet = supplierLedgerSheet.getDataRange().getValues();
	let tmp;


	formatedData.forEach(arr => {
		arr.shift();
		tmp = arr[0];
		arr.push(tmp);
		arr.shift();
	});

	Logger.log(formatedData);

	valueOfSupplierLedgerSheet.some((arr, i, self) => {
		if(self[i][2] === formatedData[0][0] - 1 ) {
			supplierLedgerSheet.insertRows(i+2,formatedData.length);
			supplierLedgerSheet.getRange(i+2,3,formatedData.length,2).setValues(formatedData);
			return true;
		}
	})
}
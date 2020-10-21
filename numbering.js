//「請求書（明細別）データ入力」で使用するスクリプト
//「【入力シート】請求書(明細別)」シートの「カスタムID未割当者抽出」ボタンを押すと実行する
function exportTriger() {
	let unNumberingData = searchUnnumberedPerson();//「【入力シート】請求書(明細別)」からカスタムIDが空欄の人を抽出する
	unNumberingData = deleateDuplicate(unNumberingData);//「unNumberingData」の重複データを削除
	let latestCustomId = findLatestCustomId(),//仕入先台帳から最新の仕入先codeを取得する
	formatedData = formatData(unNumberingData);//採番済みリストと照合するために必要情報を抜き出す
	formatedData = compareNumbered(formatedData, latestCustomId)//「採番済みリスト」と照合し、採番済みの人を除外する
	exportUunumberingData(formatedData);//「カスタムID未採番者リスト」シートに貼付を行う
}

//「請求書(明細別)データ入力」に記載された内容を削除する
function deleteTrigerInInput() {
	const inputDataInNumbering = getInputDataInNumbering();
	let lastRow = inputDataInNumbering.getLastRow(),
	lastCol     = inputDataInNumbering.getLastColumn();
	inputDataInNumbering.getRange(3, 1, lastRow, lastCol).clear();
}

function errorCountTriger() {
	const inputDataInNumbering = getInputDataInNumbering();//「【入力シート】請求書(明細別)」シートを特定
	let lastRow = inputDataInNumbering.getLastRow(),//「【入力シート】請求書(明細別)」のデータが存在する最終行を取得
	excelFunc = [];//スプレッドシート関数を入れる配列を用意

	for(let i = 3; lastRow >= i; i++) {//見出し行を省いたデータが存在する行に対して
		excelFunc.push([`=if(countif(C${i},"*テスト*")=1,"削除",if(B${i}=B${i - 1},"",if(D${i}="-","修正","")))`]);//スプレッドシート関数を挿入
	}
	inputDataInNumbering.getRange(3, 1, lastRow - 2, 1).setValues(excelFunc);//3行目以降のデータが存在する行に出力
}


//「【入力シート】請求書(明細別)」に貼り付けられたデータの中からカスタムID未採番者を特定する
//マジックナンバー
function searchUnnumberedPerson() {
	const inputData = getInputDataInNumbering(),//「【入力シート】請求書(明細別)」シートを特定
	lastRow = inputData.getLastRow(),//「【入力シート】請求書(明細別)」のデータが存在する最終行を取得
	lastCol = inputData.getLastColumn();//「【入力シート】請求書(明細別)」のデータが存在する最終列を取得
	let valueOfInputData = inputData.getRange(3, 2, lastRow, lastCol).getValues();//見出し行を除いた「【入力シート】請求書(明細別)」のデータを全件取得
	let unNumberingData = [];//カスタムIDが空欄の人を格納するための配列を用意

	for (let i = 2; valueOfInputData.length > i; i++) {//データ行数分だけ実行
		if (valueOfInputData[i][13] === '') {//カスタムIDが空欄であれば
			unNumberingData.push(valueOfInputData[i]);//unNumberingDataに行ごと挿入
		}
	}
	return unNumberingData;//unNumberingDataを戻す
}

//重複を削除する
function deleateDuplicate(unNumberingData) {
	unNumberingData = unNumberingData.filter((e, index) => {
		return !unNumberingData.some((e2, index2) =>{
			return index > index2 && e[12] == e2[12];
		});
	});
	return unNumberingData;
}

//unNumberingDataからカスタムIDが空欄の人を配列[請求元id,請求元]だけ抜き出す
function formatData(unNumberingData) {
	let tmp = [],
	formatedData = [];

	for (let i = 0; unNumberingData.length > i; i++) {//unNumberingDataの数だけ下記を実行
		tmp.push(unNumberingData[i][12]);//請求元IDを挿入する
		tmp.push(unNumberingData[i][13]);//請求元名を挿入する
    tmp.push(' ');//カスタムID用の空要素を挿入する
		formatedData.push(tmp);//tmpをformatedDataに挿入する
		tmp = [];//tmpを空にする
	}
	return formatedData;//formatedDataを戻す
}

//「採番済みリスト」と照合し、採番済みの人を除外する
function compareNumbered(formatedData, latestCustomId) {
	const numberedList = getNumberedSheet().getDataRange().getValues();//「採番済みリスト」シートデータを全件取得
	let deleateRows = [];//削除する行数を格納する配列

	for (let i = 0; formatedData.length > i; i++) {//formatedDataの数だけ下記を実行
		for (let c = 1; numberedList.length > c; c++) {//numberedListの数だけ下記を実行
			if (formatedData[i][0] === numberedList[c][0]) {
				deleateRows.push(i);
			}
		}
	}

	for (let j = 0; deleateRows.length > j; j++) {
		formatedData.splice(deleateRows[j] - j, 1)
	}
  
  formatedData.forEach((arr, i) => {
		formatedData[i][2] = latestCustomId;
		latestCustomId++;
	})

	return formatedData;
}

//「カスタムID未採番者リスト」シートに貼付を行う
function exportUunumberingData(formatedData) {
	const unnumberingPersonSheetInNumbering = getUnnumberingPersonSheetInNumbering();
	if (formatedData.length > 0) {
		unnumberingPersonSheetInNumbering.getRange(3, 1, formatedData.length, 3).setValues(formatedData);
	}
}

//「仕入先台帳」から最新の空き番号を取得する
function findLatestCustomId() {
	const supplierLedgerSheet = getSupplierLedgerSheet();
	const valueOfsupplierLedgerSheet = supplierLedgerSheet.getDataRange().getValues();
	let latestCustomId;

	valueOfsupplierLedgerSheet.some((arr, i, self) => {
		if (typeof self[i][1] == 'number' && self[i][1] === 0) {
			latestCustomId = self[i - 1][2] + 1;
			return true;
		}
	});
	return latestCustomId;
}


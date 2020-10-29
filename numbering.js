//「請求書（明細別）データ入力」で使用するスクリプト
//「【入力シート】請求書(明細別)」シートの「カスタムID未割当者抽出」ボタンを押すと実行する
function exportTriger() {
	let blankPersons = searchBlankPerson();//「【入力シート】請求書(明細別)」からカスタムIDが空欄の人を重複なしで抽出する
	let latestCustomId = findLatestCustomId();//仕入先台帳から最新の仕入先codeを取得する
	numberingData = compareNumbered(blankPersons, latestCustomId)//「採番済みリスト」と照合し、採番済みの人を除外する
	exportUunumberingData(numberingData);//「カスタムID未採番者リスト」シートに貼付を行う
}

//「請求書(明細別)データ入力」に記載された内容を削除する
function deleteTrigerInInput() {
	const inputDataInNumbering = getInputDataInNumbering();//「【入力シート】請求書(明細別)」シートを特定
	let lastRow = inputDataInNumbering.getLastRow(),//「【入力シート】請求書(明細別)」シートのデータが存在する最終行を取得
	lastCol     = inputDataInNumbering.getLastColumn();//「【入力シート】請求書(明細別)」シートのデータが存在する最終列を取得
	inputDataInNumbering.getRange(3, 1, lastRow, lastCol).clearContent();//「【入力シート】請求書(明細別)」シートの先頭2行を除いたすべてのデータを削除
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
function searchBlankPerson() {
	const valueOfInputData = getInputDataInNumbering().getDataRange().getValues();//見出し行を除いた「【入力シート】請求書(明細別)」のデータを全件取得
	let blankPersons = [],//カスタムIDが空欄の人を格納するための配列を用意
	tmp = [];//2次元配列を作成するための一時的なデータ保管変数

	for (let i = 2; valueOfInputData.length > i; i++) {//データ行数分だけ下記を実行
		if (valueOfInputData[i][14] === '') {//i行のカスタムIDが空欄であれば
			tmp.push(valueOfInputData[i][12]);//tmpにi行の請求元IDを挿入する
			tmp.push(valueOfInputData[i][13]);//tmpにi行の請求元名を挿入する
			tmp.push(' ');//tmpにカスタムID用の空要素を挿入する
			blankPersons.push(tmp);//blankPersonsにtmpを配列ごと挿入
			tmp = [];//次ループで使うのでtmpを空にする
		}
	}
	blankPersons = deleateDuplicate(blankPersons);//「blankPersons」の重複データを削除
	return blankPersons;//blankPersonsを戻す
}

//重複を削除する
//詳細仕様書を参照
function deleateDuplicate(blankPersons) {
	blankPersons = blankPersons.filter((e, index) => {
		return !blankPersons.some((e2, index2) => {
			return index > index2 && e[0] == e2[0];
		});
	});
	return blankPersons;
}

//「採番済みリスト」と照合し、採番済みの人を除外する
function compareNumbered(blankPersons, latestCustomId) {
	const numberedList = getNumberedSheet().getDataRange().getValues();//「採番済みリスト」シートデータを全件取得
	let deleateRows = [];//削除する行数を格納する配列

	//詳細仕様書を参照
	for (let i = 0; blankPersons.length > i; i++) {//blankPersonsの数だけ下記を実行
		for (let c = 1; numberedList.length > c; c++) {//numberedListの数だけ下記を実行
			if (blankPersons[i][0] === numberedList[c][0]) {
				deleateRows.push(i);
			}
		}
	}

	//詳細仕様書を参照
	for (let j = 0; deleateRows.length > j; j++) {
		blankPersons.splice(deleateRows[j] - j, 1)
	}

  //詳細仕様書を参照
  blankPersons.forEach((arr, i) => {
		blankPersons[i][2] = latestCustomId;
		latestCustomId++;
	});

	return blankPersons;
}

//「カスタムID未採番者リスト」シートに貼付を行う
function exportUunumberingData(unNumberingData) {
	const numberingPersonSheetInNumbering = getNumberingPersonSheetInNumbering();//「カスタムID未採番者リスト」シートを特定
	if (unNumberingData.length > 0) {//unNumberingDataが存在してれば
		numberingPersonSheetInNumbering.getRange(3, 1, unNumberingData.length, 3).setValues(unNumberingData);//カスタムID未採番者リスト」シートに出力
	}
}

//「仕入先台帳」から最新の空き番号を取得する
function findLatestCustomId() {
	const supplierLedgerSheet = getSupplierLedgerSheet(),//「仕入先台帳」シートを特定
	valueOfsupplierLedgerSheet = supplierLedgerSheet.getDataRange().getValues();//「仕入先台帳」シートデータを全件取得
	let latestCustomId;//最新のカスタムIDを格納する変数を宣言

	//詳細仕様書を参照
	valueOfsupplierLedgerSheet.some((arr, i, self) => {
		if (typeof self[i][1] == 'number' && self[i][1] === 0) {
			latestCustomId = self[i - 1][2] + 1;
			return true;
		}
	});
	return latestCustomId;
}


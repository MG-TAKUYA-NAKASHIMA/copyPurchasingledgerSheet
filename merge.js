// 請求書（明細別）データ入力で使用するスクリプト
//「【入力シート】請求書(明細別)」のデータと「カスタムID未採番者リスト」のデータを合わせて
//すべての人に仕入先code(カスタムID)を採番したデータを作成する
function reflectTriger() {
	initOutputDataSheet();
	let numberingList = generateNumberedList();//「カスタムID未採番者リスト」シートのシートデータを全件取得し、加工する
	exportNumberedSheet(numberingList);//「採番済みリスト」シートに出力
	exportNumberingListToSupplierLedgerSheet(numberingList);//「仕入先台帳」に出力
	let exportData = supplementExportData();//「【入力シート】請求書(明細別)」のカスタムIDを補完する
	exportSupplementsData(exportData);
}

//「【出力シート】請求書(明細別) 」シートのデータを削除し、初期化を行う
function initOutputDataSheet() {
	const outputDataInNumbering = getOutputDataInNumbering();//「【出力シート】請求書(明細別) 」シートを特定
	let lastRow = outputDataInNumbering.getLastRow();//「【出力シート】請求書(明細別) 」シートのデータが存在している最終行を取得する
	outputDataInNumbering.getRange(3, 1, lastRow, 44).clearContent();//色・サイズなどの情報を保持したまま、データだけを削除する
}

//採番された人のリストを「カスタムID未採番者リスト」から取得し、加工する
function generateNumberedList() {
	let numberingList = getNumberingPersonSheetInNumbering().getDataRange().getValues();//「カスタムID未採番者リスト」シートのデータを全件取得

	//見出しを削除する
	numberingList.shift();
	numberingList.shift();//numberingList、つまり「カスタムID未採番者リスト」シートは見出し行が先頭2行存在しているので2回shiftを行う

	return numberingList;
}

//「カスタムID未採番者」シートないしは「採番済み」シートの情報に基づいてカスタムIDを入力し、「【出力シート】請求書(明細別)」への貼付準備を行っている
function supplementExportData() {
	let valueOfInputData = getInputDataInNumbering().getDataRange().getValues(),//「【入力シート】請求書(明細別)」シートのデータを全件取得
	numberedList = getNumberedSheet().getDataRange().getValues();//「採番済みリスト」シートのデータを全件取得

	numberedList.shift();//「採番済みリスト」のシートデータの先頭行を削除

	//カスタムID未入力者にカスタムIDを割り当てる作業
	//numberedList.length < 0 の時は不要な処理
	if (numberedList.length >= 1) {//numberedListに要素が存在しており
		for (let i = 2; valueOfInputData.length > i; i++) {//valueOfInputDataの数だけ下記を実行
			for (let j = 0; numberedList.length > j; j++) {//numberedListの数だけ下記を実行
				if (valueOfInputData[i][14] == '' && valueOfInputData[i][12] === numberedList[j][0]) {//カスタムID未記入かつ請求元IDが一緒なら
					valueOfInputData[i][14] = numberedList[j][2];//カスタムIDを追加
				}
			}
		}
	}
	//見出し行の削除
	valueOfInputData.shift();
	valueOfInputData.shift();

	//判別列(A列)の削除
	valueOfInputData.forEach(arr => {
			arr.shift();
		});

	return valueOfInputData;
}


//仕入先台帳へ反映
function exportNumberingListToSupplierLedgerSheet(numberingList) {
	const supplierLedgerSheet = getSupplierLedgerSheet();//「仕入先台帳」シートを特定
	let valueOfSupplierLedgerSheet = supplierLedgerSheet.getDataRange().getValues(),//「仕入先台帳」シートを全件取得する
	tmp;//配列の要素入れ替え用の一時変数を宣言

	numberingList.forEach(arr => {
		arr.shift();//請求元IDを削除
		tmp = arr[0];//表示名を最終indexに持っていくために退避
		arr.push(tmp);//表示名を最終indexにpush
		arr.shift();//先頭の表示名を削除
	});

	valueOfSupplierLedgerSheet.some((arr, i, self) => {
		if (numberingList.length > 0 && self[i][2] === numberingList[0][0] - 1) {//numberingListが1以上かつ、仕入先台帳のBiセルが採番された番号 - 1と同一の場合
			supplierLedgerSheet.insertRows(i + 2, numberingList.length);//numberingListの数だけの行数を挿入する
			supplierLedgerSheet.getRange(i + 2, 3, numberingList.length, 2).setValues(numberingList);//numberingListを出力する
			return true;
		}
	});
}

//「採番済みリスト」シートへの書き込み
function exportNumberedSheet(numberingList) {
	const numberedSheet = getNumberedSheet(),//「採番済みリスト」シートを特定
		lastRow = numberedSheet.getLastRow(),//「採番済みリスト」シートデータが存在する最終行を取得
		numberingListSheet = getNumberingPersonSheetInNumbering();//「カスタムID未採番者リスト」シートを特定

	if (numberingList.length >= 1) {//numberingListにデータが存在していれば
		numberedSheet.getRange(lastRow + 1, 1, numberingList.length, numberingList[0].length).setValues(numberingList);//「採番済みリスト」の最終行 + 1行の位置にnuberingListを出力
		numberingListSheet.getRange(3, 1, numberingList.length, 3).clearContent();//「カスタムID未採番者リスト」シートのデータのみを削除
	}
}

//「【出力シート】請求書(明細別)」シートへ出力
function exportSupplementsData(exportData) {
	const outputDataInNumbering = getOutputDataInNumbering();//「【出力シート】請求書(明細別) 」シートを特定
	outputDataInNumbering.getRange(3, 1, exportData.length, exportData[0].length).setValues(exportData);//「【出力シート】請求書(明細別)」シートへ出力
}

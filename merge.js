// 請求書（明細別）データ入力で使用するスクリプト
//「【入力シート】請求書(明細別)」のデータと「カスタムID未採番者リスト」のデータを合わせて
//すべての人に仕入先code(カスタムID)を採番したデータを作成する
function reflectTriger() {
	initOutputDataSheet();
	let numberingList = generateNumberedList(),//採番された人のリストを「カスタムID未採番者リスト」と「採番済みリスト」から取得する
		exportData = supplementExportData(numberingList);//「【入力シート】請求書(明細別)」のカスタムIDを補完する
	exportSupplementsData(exportData);
}

//「【出力シート】請求書(明細別) 」シートのデータを削除し、初期化を行う
function initOutputDataSheet() {
	const outputDataInNumbering = getOutputDataInNumbering();//「【出力シート】請求書(明細別) 」シートを特定
	let lastRow = outputDataInNumbering.getLastRow();//「【出力シート】請求書(明細別) 」シートのデータが存在している最終行を取得する
	outputDataInNumbering.getRange(3, 1, lastRow, 44).clearContent();//色・サイズなどの情報を保持したまま、データだけを削除する
}

//採番された人のリストを「カスタムID未採番者リスト」と「採番済みリスト」から取得する
function generateNumberedList() {
	let numberingList = getNumberingPersonSheetInNumbering().getDataRange().getValues(),//「カスタムID未採番者リスト」シートのデータを全件取得
		numberedList = getNumberedSheet().getDataRange().getValues();//「採番済みリスト」シートのデータを全件取得

	//見出しを削除する
	numberingList.shift();
	numberingList.shift();//numberingList、つまり「カスタムID未採番者リスト」シートは見出し行が戦闘2行存在しているので2回shiftを行い
	numberedList.shift();//numberedList、「採番済みリスト」シートには見出し行が1行しか存在していないので1回shiftを行う

	exportNumberingListToSupplierLedgerSheet(numberingList);//「仕入先台帳」に出力
	exportNumberedSheet(numberingList);//「採番済みリスト」シートに出力

	//以下で「カスタムID未採番者リスト」由来のリストと「採番済みリスト」由来リストをつなげている
	for (let i = 0; numberedList.length > i; i++) {
		numberingList.push(numberedList[i]);
	}
		return numberingList;
}

//「カスタムID未採番者」シートないしは「採番済み」シートの情報に基づいてカスタムIDを入力し、「【出力シート】請求書(明細別)」への貼付準備を行っている
function supplementExportData(numberingList) {
	let valueOfInputData = getInputDataInNumbering().getDataRange().getValues();//「【入力シート】請求書(明細別)」シートのデータを全件取得

	//カスタムID未入力者にカスタムIDを割り当てる作業
	//numberingList.length < 0 の時は不要な処理
	if (numberingList.length >= 1) {//numberingListに要素が存在しており
		for (let i = 2; valueOfInputData.length > i; i++) {//valueOfInputDataの数だけ下記を実行
			for (let j = 0; numberingList.length > j; j++) {//numberingListの数だけ下記を実行
				if (valueOfInputData[i][14] == '' && valueOfInputData[i][12] === numberingList[j][0]) {//カスタムID未記入かつ請求元IDが一緒なら
					valueOfInputData[i][14] = numberingList[j][2];//カスタムIDを追加
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
function exportNumberedSheet() {
	const numberedSheet = getNumberedSheet(),//「採番済みリスト」シートを特定
		lastRow = numberedSheet.getLastRow(),//「採番済みリスト」シートデータが存在する最終行を取得
		numberingListSheet = getNumberingPersonSheetInNumbering();//「カスタムID未採番者リスト」シートを特定
	let numberingList = numberingListSheet.getDataRange().getValues();//「カスタムID未採番者リスト」シートデータを全件取得

	//「カスタムID未採番者リスト」の見出し行を2行分削除
	numberingList.shift();
	numberingList.shift();
	
	
	if (numberingList.length >= 1) {//numberingListにデータが存在していれば
		numberedSheet.getRange(lastRow + 1, 1, numberingList.length, 3).setValues(numberingList);//「採番済みリスト」の最終行 + 1行の位置にnuberingListを出力
		numberingListSheet.getRange(3, 1, numberingList.length, 3).clearContent();//「カスタムID未採番者リスト」シートのデータのみを削除
	}
}

//「【出力シート】請求書(明細別)」シートへ出力
function exportSupplementsData(exportData) {
	const outputDataInNumbering = getOutputDataInNumbering();//「【出力シート】請求書(明細別) 」シートを特定
	outputDataInNumbering.getRange(3, 1, exportData.length, exportData[0].length).setValues(exportData);//「【出力シート】請求書(明細別)」シートへ出力
}

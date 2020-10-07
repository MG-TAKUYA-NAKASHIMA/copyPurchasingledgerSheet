// 請求書（明細別）データ入力で使用するスクリプト

//「【入力シート】請求書(明細別)」のデータと
//「カスタムID未採番者リスト」のデータを合わせて
//すべての人に仕入先code(カスタムID)を採番したデータを作成する
//3,1,44はマジックナンバーなので修正
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

//unNumberungList=「カスタムID未採番者リスト」に出力されていたもの
//numberedList=「採番済みリスト」に出力されていたもの
//上記の2つの2次元配列の構成の構成がおかしいが時間的問題で修正を行っていない
//→generateMergeData関数でのforとifの入れ子につながっている
//unNumberungList = [カスタムID,pasture表示氏名]
//numberedList= [請求元ID,pasture表示氏名,カスタムID]
//また配列の変数名を修正したい。
for(let i = 0; numberedList.length > i; i++) {
	unNumberingList.push(numberedList[i]);
}
  Logger.log(unNumberingList);
return unNumberingList;

}

//「カスタムID未採番者」シートと「採番済み」シートという2つのシートから採番されていない人を読み込んでいる
//forとifの入れ子は改善しなければいけないが、時間的な制約により修正できていない(原因は上記関数のコメントを参照)
//なので配列要素数でunNumberungListに由来するのか、numberedListに由来するのかを判別させている。
function generateMergeData(unNumberingList) {
	let valueOfInputData = getInputDataInNumbering().getDataRange().getValues();

	for (let i = 2; valueOfInputData.length > i; i++) {
		if (valueOfInputData[i][13] == '') {
			for (let j = 0; unNumberingList.length > j; j++) {
              if(unNumberingList[j].length == 2 ){
               if (valueOfInputData[i][12] == unNumberingList[j][1]) {
					valueOfInputData[i][13] = unNumberingList[j][0];
				}
              }else if(unNumberingList[j].length == 3){
				if (valueOfInputData[i][12] == unNumberingList[j][1]) {
					valueOfInputData[i][13] = unNumberingList[j][2];
				}
              }
			}
		}
	}

	valueOfInputData.shift().shift();
	return valueOfInputData;
}

//3,1はマジックナンバーなので修正
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


	//i + 2はマジックナンバーなので変数を宣言してあげるように修正 
	valueOfSupplierLedgerSheet.some((arr, i, self) => {
		if(unNumberingList.length > 0 && self[i][2] === unNumberingList[0][0] - 1 ) {
			supplierLedgerSheet.insertRows(i+2,unNumberingList.length);
			supplierLedgerSheet.getRange(i+2,3,unNumberingList.length,2).setValues(unNumberingList);
			return true;
		}
	})
}


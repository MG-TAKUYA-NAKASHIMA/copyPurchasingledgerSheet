function detectErrorTriger() {
	const inputDataInNumbering = getInputDataInNumbering();//「【入力シート】請求書(明細別)」シートを特定
	let lastRow = inputDataInNumbering.getLastRow(),//「【入力シート】請求書(明細別)」のデータが存在する最終行を取得
	excelFunc = [];//スプレッドシート関数を入れる配列を用意

	for(let i = 3; lastRow >= i; i++) {//見出し行を省いたデータが存在する行に対して
		excelFunc.push([`=if(A${i - 1}="削除",if(D${i}="-","削除",""),if(countif(E${i},"*テスト*")=1,"削除",if(B${i}=B${i - 1},"",if(D${i}="-","修正",""))))`]);//スプレッドシート関数を挿入
	}
  inputDataInNumbering.getRange(3, 1, lastRow - 2, 1).setValues(excelFunc);//3行目以降のデータが存在する行に出力
  deleteRow();
}

function deleteRow() {
  const inputDataInNumbering = getInputDataInNumbering();
  let valueOfInputDataInNumbering = inputDataInNumbering.getDataRange().getValues(),
  deleteRows = [],//削除対象行を格納する変数を宣言
	excelFunc = [];//スプレッドシート関数を入れる配列を用意

  valueOfInputDataInNumbering.splice(0,2);//「【入力シート】請求書(明細別)」シートデータの見出し2行を削除
  valueOfInputDataInNumbering.forEach((arr,i) => {
    if(arr[0] === "削除") {
      deleteRows.push(i + 3);
    }
  });

  for (let j = 0; deleteRows.length > j; j++) {//deleteRowsの数だけ下記を実行
    inputDataInNumbering.deleteRows(deleteRows[j] - j);//削除する
  }

  let lastRow = inputDataInNumbering.getLastRow();//「【入力シート】請求書(明細別)」のデータが存在する最終行を取得

  for(let k = 3; lastRow >= k; k++) {//見出し行を省いたデータが存在する行に対して
		excelFunc.push([`=if(A${k - 1}="削除",if(D${k}="-","削除",""),if(countif(E${k},"*テスト*")=1,"削除",if(B${k}=B${k - 1},"",if(D${k}="-","修正",""))))`]);//スプレッドシート関数を挿入
	}
  inputDataInNumbering.getRange(3, 1, lastRow - 2, 1).setValues(excelFunc);//3行目以降のデータが存在する行に出力
}
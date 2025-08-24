// メインの実行関数
function processDailyReport() {
    try {
        // 列削除を個別に実行し、完了を確認
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const columnLetters = ['L', 'K', 'J', 'I', 'G'];  // 順序を修正
        const columnIndices = columnLetters.map(letter =>
            letter.charCodeAt(0) - 'A'.charCodeAt(0) + 1
        );

        // 列削除を1つずつ実行
        columnIndices.sort((a, b) => b - a).forEach(columnIndex => {
            try {
                sheet.deleteColumn(columnIndex);
                SpreadsheetApp.flush(); // 各削除操作後に確実に適用
            } catch (deleteError) {
                Logger.log('列削除でエラー: ' + deleteError.toString());
                // 個別の削除エラーは全体の処理を止めない
            }
        });

        // 他の処理を実行
        Utilities.sleep(500); // 列削除の完了を待機

        replaceBoolean();
        clearBasedOnCondition();
        addSequenceNumbers();
        formatAndCombineText();
        insertHeaders();
        setColumnWidthsAndMerge();
        setBorders();
        addSummary();

        SpreadsheetApp.flush();
    } catch (error) {
        Logger.log('エラーが発生しました: ' + error.toString());
        throw error;
    }
}

// 以下の共通ユーティリティ関数と他の関数は同じ
function getValidDataRange() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();
    const cColumn = sheet.getRange('C3:C' + lastRow).getValues();
    let lastValidRow = 3;

    for (let i = 0; i < cColumn.length; i++) {
        if (cColumn[i][0] && cColumn[i][0].toString().trim() !== '') {
            lastValidRow = i + 3;
        }
    }

    return {
        sheet: sheet,
        lastRow: lastRow,
        lastValidRow: lastValidRow
    };
}

function insertHeaders() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    const headers = [
        '弁当',
        '氏名',
        '欠席',
        '当キャ',
        '出勤',
        '退勤',
        '作業内容',
        '日報'
    ];

    sheet.getRange('B2:I2').setValues([headers]);
}

function replaceBoolean() {
    const { sheet, lastValidRow } = getValidDataRange();
    const dataRange = sheet.getRange(1, 1, lastValidRow, sheet.getLastColumn());
    const values = dataRange.getValues();

    values.forEach((row, i) => {
        row.forEach((cell, j) => {
            if (cell === true) values[i][j] = '○';
            else if (cell === false) values[i][j] = '';
        });
    });

    dataRange.setValues(values);
}

function clearBasedOnCondition() {
    const { sheet, lastValidRow } = getValidDataRange();
    const rangeD = sheet.getRange('D3:D' + lastValidRow);
    const rangeE = sheet.getRange('E3:E' + lastValidRow);
    const valuesD = rangeD.getValues();
    const valuesE = rangeE.getValues();

    valuesD.forEach((value, i) => {
        if (value[0] === '○' || valuesE[i][0] === '○') {
            sheet.getRange(i + 3, 6, 1, 2).clearContent();
        }
    });
}

function addSequenceNumbers() {
    const { sheet, lastValidRow } = getValidDataRange();
    let sequenceNumber = 1;

    for (let row = 3; row <= lastValidRow; row++) {
        sheet.getRange(row, 1).setValue(sequenceNumber++);
    }
}

function formatAndCombineText() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();

    // 3行目から処理開始
    for (let row = 3; row <= lastRow; row++) {
        // H列からO列までの値を取得
        const range = sheet.getRange(row, 8, 1, 8); // H=8, O=15
        const values = range.getValues()[0];

        // 数値を除外し、文字列のみを抽出
        const texts = values.filter(value =>
            typeof value === 'string' && value.toString().trim() !== ''
        );

        // 重複を除去
        const uniqueTexts = [...new Set(texts)];

        // 文字列を「、」で結合
        const combinedText = uniqueTexts.join('、');

        // H列に結合した文字列を設定
        sheet.getRange(row, 8).setValue(combinedText);
    }

    // I列からO列を削除（結合後に一括で削除）
    for (let col = 15; col >= 9; col--) {  // O列(15)からI列(9)まで降順で削除
        sheet.deleteColumn(col);
    }
}

function setColumnWidthsAndMerge() {
    const { sheet, lastValidRow } = getValidDataRange();

    // 列幅を固定値に設定
    sheet.setColumnWidth(7, 48);  // G列
    sheet.setColumnWidth(8, 135); // H列
    sheet.setColumnWidth(9, 240); // I列

    // H列とI列のテキスト折り返し設定
    sheet.getRange(2, 8, lastValidRow - 1, 2).setWrap(true);  // 2行目からlastValidRowまでのH列とI列
}

function setBorders() {
    const { sheet, lastValidRow } = getValidDataRange();

    // A列からI列、2行目から最終有効行までの範囲を取得（範囲を修正）
    const range = sheet.getRange(2, 1, lastValidRow - 1, 9); // 9はI列

    // すべての罫線を設定
    range.setBorder(
        true, // top
        true, // left
        true, // bottom
        true, // right
        true, // vertical
        true, // horizontal
        'black', // color
        SpreadsheetApp.BorderStyle.SOLID // style
    );
}

function addSummary() {
    const { sheet, lastValidRow } = getValidDataRange();
    
    // 集計開始行（有効データ最終行 + 2）
    const startRow = lastValidRow + 2;
    const startCol = 6;  // F列から開始
    
    // 項目と数式を設定
    const items = [
      {
        label: '午前のみ利用',
        formula: `=COUNTIFS(F3:F${lastValidRow}, "<>", G3:G${lastValidRow}, "<13:00")`,
      },
      {
        label: '午後のみ利用',
        formula: `=COUNTIFS(F3:F${lastValidRow}, ">=13:00")`,
      },
      {
        label: '午前～午後利用',
        formula: `=COUNTIFS(F3:F${lastValidRow}, "<11:30", G3:G${lastValidRow}, ">=13:30")`,
      },
      {
        label: '合計',
        formula: `=SUM(${sheet.getRange(startRow, startCol + 2, 3).getA1Notation()})`,
      },
      {
        label: '',  // 空行
        formula: '',
      },
      {
        label: '当日キャンセル',
        formula: `=COUNTIF(E3:E${lastValidRow}, "○")`,
      },
      {
        label: '欠席',
        formula: `=COUNTIF(D3:D${lastValidRow}, "○")`,
      }
    ];
    
    // データを設定
    items.forEach((item, index) => {
      const row = startRow + index;
      
      // 項目名（1列目）
      sheet.getRange(row, startCol).setValue(item.label);
      
      // 2列目は空欄のまま
      
      // 数式（3列目）
      if (item.formula) {
        sheet.getRange(row, startCol + 2).setFormula(item.formula);
      }
      
      // 単位「人」（4列目）
      if (item.label && item.label !== '') {
        sheet.getRange(row, startCol + 3).setValue('人');
      }
    });
    
    // 合計行の下に罫線を引く
    const borderRow = startRow + 3;  // 合計行
    sheet.getRange(borderRow, startCol, 1, 4).setBorder(  // 4列分の範囲
      false, // top
      false, // left
      true,  // bottom
      false, // right
      false, // vertical
      false, // horizontal
      'black',
      SpreadsheetApp.BorderStyle.SOLID
    );
  }
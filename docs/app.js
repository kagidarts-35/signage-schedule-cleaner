const input = document.querySelector('#file');
const dropzone = document.querySelector('#dropzone');
const processButton = document.querySelector('#process');
const errorBox = document.querySelector('#error');
const resultBox = document.querySelector('#result');
let selectedFile = null;
let downloadUrl = null;

function showError(message) {
  errorBox.textContent = message;
  errorBox.hidden = false;
}

function selectFile(file) {
  if (!file) return;
  if (!/\.xlsx$/i.test(file.name)) return showError('Excelファイル（.xlsx）を選択してください。');
  if (file.size > 30 * 1024 * 1024) return showError('30MB以下のExcelファイルを選択してください。');
  if (downloadUrl) URL.revokeObjectURL(downloadUrl);
  selectedFile = file;
  downloadUrl = null;
  errorBox.hidden = true;
  resultBox.hidden = true;
  dropzone.classList.add('selected');
  document.querySelector('#file-name').textContent = file.name;
  document.querySelector('#file-help').textContent = `${Math.round(file.size / 1024)} KB・クリックして変更`;
  processButton.disabled = false;
}

dropzone.addEventListener('click', () => input.click());
dropzone.addEventListener('keydown', (event) => {
  if (event.key === 'Enter' || event.key === ' ') input.click();
});
dropzone.addEventListener('dragover', (event) => event.preventDefault());
dropzone.addEventListener('drop', (event) => {
  event.preventDefault();
  selectFile(event.dataTransfer.files[0]);
});
input.addEventListener('change', () => selectFile(input.files[0]));

function textOf(value) {
  if (typeof value === 'string') return value;
  if (value && Array.isArray(value.richText)) return value.richText.map((item) => item.text).join('');
  return value == null ? '' : String(value);
}

processButton.addEventListener('click', async () => {
  if (!selectedFile) return;
  processButton.disabled = true;
  processButton.textContent = '修正しています…';
  errorBox.hidden = true;
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await selectedFile.arrayBuffer());
    let replacements = 0;
    let deletedCount = 0;
    const deleteHeaders = ['担当', '終了日', '備考'];

    workbook.eachSheet((sheet) => {
      const deleteColumns = [];
      let programColumn = null;
      for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 15); rowNumber += 1) {
        sheet.getRow(rowNumber).eachCell({ includeEmpty: false }, (cell, columnNumber) => {
          const text = textOf(cell.value).trim();
          if (deleteHeaders.includes(text)) deleteColumns.push(columnNumber);
          if (text.includes('会社名・番組名')) programColumn = columnNumber;
        });
      }
      if (programColumn !== null) {
        for (let rowNumber = 4; rowNumber <= sheet.rowCount; rowNumber += 1) {
          const cell = sheet.getCell(rowNumber, programColumn);
          if (typeof cell.value === 'string' && cell.value.includes('グレイド枠')) {
            cell.value = cell.value.replaceAll('グレイド枠', '');
            replacements += 1;
          }
        }
      }
      [...new Set(deleteColumns)].sort((a, b) => b - a).forEach((column) => {
        sheet.spliceColumns(column, 1);
        deletedCount += 1;
      });
      for (let row = 1; row <= sheet.rowCount; row += 1) {
        for (let column = 1; column <= sheet.columnCount; column += 1) {
          const cell = sheet.getCell(row, column);
          if (cell.value !== null || cell.style) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
          }
        }
      }
    });

    const output = await workbook.xlsx.writeBuffer();
    downloadUrl = URL.createObjectURL(new Blob([output], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }));
    const download = document.querySelector('#download');
    download.href = downloadUrl;
    download.download = `${selectedFile.name.replace(/\.xlsx$/i, '')}_クライアント用.xlsx`;
    document.querySelector('#summary').textContent = `${deletedCount}列を削除・「グレイド枠」を${replacements}セルから削除・背景を白に変更`;
    resultBox.hidden = false;
  } catch (error) {
    console.error(error);
    showError('処理できませんでした。Excelで開ける通常の .xlsx ファイルか確認してください。');
  } finally {
    processButton.disabled = false;
    processButton.innerHTML = '番組表を修正する <span>→</span>';
  }
});

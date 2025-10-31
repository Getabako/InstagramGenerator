// 画像のベースURL
const BASE_URL = 'https://images.if-juku.net';

/**
 * スプレッドシート起動時にカスタムメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('画像リンク')
    .addItem('画像リンクを挿入', 'insertImageLinks')
    .addToUi();
}

/**
 * D列が空の場合、画像リンクを挿入
 * 4つの画像+サンクスメッセージ画像の5枚セットをカンマ区切りでD列に配置
 */
function insertImageLinks() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();

  // A列とD列の値をチェック（2行目以降）
  if (lastRow > 1) {
    const aColumnValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const dColumnValues = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
    const hasAContent = aColumnValues.some(row => row[0] !== '');
    const hasDContent = dColumnValues.some(row => row[0] !== '');

    if (hasAContent || hasDContent) {
      ui.alert('A列またはD列にすでにデータが入力されています。\n処理を中止します。');
      return;
    }
  }

  try {
    // ユーザー入力を受け取る
    const folderNameResponse = ui.prompt(
      '画像リンク生成',
      'フォルダ名を入力してください\n（例：tech_post_2025_10）',
      ui.ButtonSet.OK_CANCEL
    );

    if (folderNameResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    const folderName = folderNameResponse.getResponseText().trim();

    if (!folderName) {
      ui.alert('フォルダ名が入力されていません。');
      return;
    }

    // 画像枚数の入力
    const imageCountResponse = ui.prompt(
      '画像リンク生成',
      '画像の枚数を入力してください（4の倍数）\n（例：8）',
      ui.ButtonSet.OK_CANCEL
    );

    if (imageCountResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    const imageCount = parseInt(imageCountResponse.getResponseText().trim());

    if (isNaN(imageCount) || imageCount <= 0) {
      ui.alert('有効な数値を入力してください。');
      return;
    }

    if (imageCount % 4 !== 0) {
      ui.alert('画像枚数は4の倍数である必要があります。\n入力された枚数：' + imageCount);
      return;
    }

    // サンクスメッセージ画像名の入力
    const thanksImageResponse = ui.prompt(
      '画像リンク生成',
      'サンクスメッセージの画像名を入力してください\n（例：iftech_thanks.png）',
      ui.ButtonSet.OK_CANCEL
    );

    if (thanksImageResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    const thanksImageName = thanksImageResponse.getResponseText().trim();

    if (!thanksImageName) {
      ui.alert('サンクスメッセージ画像名が入力されていません。');
      return;
    }

    // 画像リンクを生成
    const imageLinks = [];
    for (let i = 1; i <= imageCount; i++) {
      const imageNumber = String(i).padStart(3, '0'); // 001, 002, ...
      const imageUrl = `${BASE_URL}/${folderName}/${imageNumber}.png`;
      imageLinks.push(imageUrl);
    }

    // サンクスメッセージ画像のリンクを生成
    const thanksImageLink = `${BASE_URL}/${folderName}/${thanksImageName}`;

    // 4つずつのグループに分割し、各グループにサンクスメッセージ画像を追加してカンマ区切りで連結
    const rowData = [];
    for (let i = 0; i < imageLinks.length; i += 4) {
      const imageSet = [
        imageLinks[i],
        imageLinks[i + 1],
        imageLinks[i + 2],
        imageLinks[i + 3],
        thanksImageLink
      ];
      const combinedLinks = imageSet.join(',');
      rowData.push([combinedLinks]);
    }

    // 日付データを作成（明日から開始し、毎日18:00に設定）
    const dateData = [];
    const startDate = new Date();
    startDate.setDate(startDate.getDate() + 1); // 明日に設定
    startDate.setHours(18, 0, 0, 0); // 時刻を18:00:00に設定

    for (let i = 0; i < rowData.length; i++) {
      const currentDate = new Date(startDate);
      currentDate.setDate(startDate.getDate() + i); // i日後の日付

      // 日付を「YYYY-MM-DD HH:mm」形式の文字列にフォーマット（Publer推奨形式）
      const year = currentDate.getFullYear();
      const month = String(currentDate.getMonth() + 1).padStart(2, '0');
      const day = String(currentDate.getDate()).padStart(2, '0');
      const hours = String(currentDate.getHours()).padStart(2, '0');
      const minutes = String(currentDate.getMinutes()).padStart(2, '0');
      const formattedDate = `${year}-${month}-${day} ${hours}:${minutes}`;

      dateData.push([formattedDate]);
    }

    // 1行目にヘッダーを追加
    sheet.getRange(1, 1, 1, 4).setValues([['Date', 'Text', 'Link(s)', 'Media URL(s)']]);

    // A列に日付、D列に画像リンクを挿入（2行目から）
    sheet.getRange(2, 1, dateData.length, 1).setValues(dateData);
    sheet.getRange(2, 4, rowData.length, 1).setValues(rowData);

    const groupCount = rowData.length;
    ui.alert(
      `${imageCount}件の画像リンクを挿入しました。\n` +
      `フォルダ名：${folderName}\n` +
      `グループ数：${groupCount}\n` +
      `各グループ：4画像+サンクスメッセージ画像`
    );

  } catch (error) {
    ui.alert(`エラーが発生しました: ${error.message}`);
  }
}

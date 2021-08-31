function extractNutsRecepi() {
  // スプレッドシートを取得
  const ss = SpreadsheetApp.getActiveSheet();

  // 入力値をリセット
  const lastDataRow = ss.getRange(6 ,1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow()
  ss.getRange(6 ,1, lastDataRow-5, 1).clear();

  // ナッツ名を取得
  const searchWord = ss.getRange(3, 1).getValue();
  const encodeSearchWord = encodeURI(searchWord);

  // 検索結果を入れる配列
  let recipeList = [];

  // 10ページ目までデータを取得
  for (j=1; j<11; j++) {//10ページ目までの検索結果を取得
    const pageNum = j;
    const url = 'https://cookpad.com/search/'+ encodeSearchWord +'?order=date&page=' + pageNum;
    const options = {muteHttpExceptions:true};
    let responseText = UrlFetchApp.fetch(url, options).getContentText();

    let _recipeList = responseText.match(/<a class="recipe-title font13.*<\/a>/g);
    // Logger.log(_recipeList);
    // _recipeList[0].replace(/<a.*recipe\/\d{4,7}/,);
    if (j>1 && _recipeList === null) {// 検索ページが無くなった時点で処理を終了する
      continue;
    } else if (j===1 && _recipeList === null){// 検索結果が取得できなかった場合、処理を終了する
      return Browser.msgBox("レシピが存在しません");
    } else {
      for (i=0; i<_recipeList.length; i++) {
        let recipeNames = _recipeList[i].match(/recipe\/\d{4,7}">.*</g);
        let splitRecipe = recipeNames[0].split(/\/|">|</);
        let recipeName = splitRecipe[2];
        let recipeUrl = "https://cookpad.com/recipe/" + splitRecipe[1];
        let setFormula = `=HYPERLINK("${recipeUrl}","${recipeName}")`;
        recipeList.push([setFormula]);
      }
    }
    Utilities.sleep(1500);//サーバーに負担をかけないように。1.5秒のスリープタイムを取る
  }
  ss.getRange(6, 1, recipeList.length, 1).setValues(recipeList);
  Browser.msgBox("検索が完了しました！");
}

function todayNutsRecepi() {
  const ss = SpreadsheetApp.getActiveSheet();
  const lastDataRow = ss.getRange(5 ,1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow()
  const nutsRecipes = ss.getRange(6 ,1, lastDataRow-5, 1).getFormulas();

  // リスト範囲内のランダムな数を生成する
  const randomNum = Math.floor(Math.random()*nutsRecipes.length);
  ss.getRange(4, 5).setValue(nutsRecipes[randomNum][0]).setFontWeight('bold').setFontSize(12);
}

function extractNutsName() {
  // ナッツの名前を小島屋HPから取得
  const nutsUrl = 'https://www.kojima-ya.com/c/all/nut';
  let responseText = UrlFetchApp.fetch(nutsUrl).getContentText();

  let nutsList = [];
  let _nutsList = responseText.match(/<h2 class="fs\-c\-subgroupList__label_txt">.*<\/h2>/g);
  for (i=0; i<_nutsList.length; i++) {
    let nutsName = _nutsList[i].replace(/<h2 class="fs\-c\-subgroupList__label_txt">|<\/h2>/g, "");
    nutsList.push(nutsName);
  }

  // スプレッドシートにリストとして挿入
  const ss = SpreadsheetApp.getActiveSheet();
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(nutsList).build();
  ss.getRange(3, 1).setDataValidation(rule);
  ss.getRange(4, 1).setDataValidation(rule);
}

function test() {
  let _a = '<a class="recipe-title font13 " id="recipe_title_6643532" href="/recipe/6643532">豚バラのキャベツロールです。</a>';
  let a = _a.replace(/<a.*recipe\/|<\/a>/g, "");
  Logger.log(a);
  let recipeData = a.split('">');
  Logger.log(recipeData);
}
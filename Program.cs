using GrapeCity.Documents.Excel;

//Workbook.SetLicenseKey("");

// 新規ワークブックの作成
Workbook workbook = new();

// 外部参照式を設定
workbook.Worksheets[0].Range["E5"].Formula = @"='[workbook1.xlsx]Sheet1'!B2";

// 外部ワークブックを読み込み
Workbook externalworkbook = new();
externalworkbook.Open("workbook1.xlsx");

// 外部参照を更新
foreach (var item in workbook.GetExcelLinkSources())
{
    workbook.UpdateExcelLink(item, externalworkbook);
}

// EXCELファイルに保存
workbook.Save("crossworkbookformula.xlsx");

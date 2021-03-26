namespace GrapeCity.Documents.Excel.Examples.Features.Formulas.CrossWorkbookFormula
{
    class Program
    {
        static void Main(string[] args)
        {
            //Workbook.SetLicenseKey("");

            // 新規ワークブックの作成
            var workbook = new Workbook();

            // 外部参照式を設定
            workbook.Worksheets[0].Range["E5"].Formula = @"='[workbook1.xlsx]Sheet1'!B2";

            // 外部ワークブックを読み込み
            var externalworkbook = new Workbook();
            externalworkbook.Open("workbook1.xlsx");

            // 外部参照を更新
            foreach (var item in workbook.GetExcelLinkSources())
            {
                workbook.UpdateExcelLink(item, externalworkbook);
            }

            // EXCELファイル（.xlsx）に保存
            workbook.Save("crossworkbookformula.xlsx");

            // PDFファイル（.pdf）に保存
            workbook.Save("crossworkbookformula.pdf");
        }
    }
}
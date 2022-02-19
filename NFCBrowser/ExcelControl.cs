using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Windows;

namespace NFCBrowser
{
    class ExcelControl
    {
        const int icoca_start_col = 1;
        const int icoca_start_row = 5;
        const int name_col = 8;
        const int name_row = 8;
        const int year_col = 8;
        const int year_row = 5;
        const int month_col = 10;
        const int month_row = 5;
        const int day_col = 12;
        const int day_row = 5;
        const int customer_col = 10;
        const int customer_row = 13;
        const int total_price_col = 10;
        const int total_price_row = 19;
        IWorkbook book;
        string file_path = "";
        public ExcelControl()
        {
            // デフォルト値
            
        }
        public IWorkbook CreateNewBook(string filePath)
        {
            var extension = Path.GetExtension(filePath);

            // HSSF => Microsoft Excel(xls形式)(excel 97-2003)
            // XSSF => Office Open XML Workbook形式(xlsx形式)(excel 2007以降)
            if (extension == ".xls")
            {
                book = new HSSFWorkbook();
            }
            else if (extension == ".xlsx")
            {
                book = new XSSFWorkbook();
            }
            else
            {
                throw new ApplicationException("CreateNewBook: invalid extension");
            }

            return book;
        }
        /// <summary>
        /// エクセルファイル開く
        /// </summary>
        /// <param name="filePath"></param>
        public void OpenWorkbook(string filePath)
        {
            book = WorkbookFactory.Create(filePath);
            file_path = filePath;
        }
        /// <summary>
        /// データをセットする
        /// </summary>
        /// <param name="data"></param>
        public void SetICOCA(DataRow[] rows)
        {
            ISheet sheet = book.GetSheet("通常使用");
            for (int i = 0; i < rows.Length; i++)
            {
                writeCellValue(sheet, icoca_start_col, icoca_start_row + i, "20" + rows[i]["Date"].ToString());
                writeCellValue(sheet, icoca_start_col + 1, icoca_start_row + i, rows[i]["IriSen"].ToString()+":"+ rows[i]["Iri"].ToString());
                writeCellValue(sheet, icoca_start_col + 2, icoca_start_row + i, rows[i]["DeSen"].ToString() + ":" + rows[i]["De"].ToString());
                WriteCell(sheet, icoca_start_col + 3, icoca_start_row + i, double.Parse(rows[i]["Shiharai"].ToString()));
                WriteCell(sheet, icoca_start_col + 4, icoca_start_row + i, double.Parse(rows[i]["Zandaka"].ToString()));

            }
        }
        public void SetOthers(Dictionary<string,string> dc)
        {
            ISheet sheet = book.GetSheet("通常使用");
            for (int i = 0; i < dc.Count; i++)
            {
                writeCellValue(sheet, name_col, name_row, dc["name"]);
                writeCellValue(sheet, customer_col, customer_row, dc["cust"]);
                WriteCell(sheet, year_col, year_row, double.Parse(dc["year"].ToString()));
                WriteCell(sheet, month_col, month_row, double.Parse(dc["month"].ToString()));
                WriteCell(sheet, day_col, day_row, double.Parse(dc["day"].ToString()));
                WriteCell(sheet, total_price_col, total_price_row, double.Parse(dc["total"].ToString()));


            }
        }
        public void CloseWorkbook()
        {
            // ファイルあったら削除
            string FilePath = "ICOCA利用明細書.xlsx";
            FileInfo fi = new FileInfo(FilePath);
            if (fi.Exists)
            {
                if ((fi.Attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                {
                    fi.Attributes = FileAttributes.Normal;
                }
            }
            fi.Delete();


            // ファイル保存
            using (var fs = new FileStream("ICOCA利用明細書.xlsx", FileMode.Create))
            {
                book.Write(fs);

            }
            book.Close();
        }

        /// <summary>
        /// 文字列の書き込み
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="idxColumn"></param>
        /// <param name="idxRow"></param>
        /// <param name="value"></param>
        private void writeCellValue(ISheet sheet, int idxColumn, int idxRow, string value)
        {
            var row = sheet.GetRow(idxRow) ?? sheet.CreateRow(idxRow); //指定した行を取得できない時はエラーとならないよう新規作成している
            var cell = row.GetCell(idxColumn) ?? row.CreateCell(idxColumn); //一行上の処理の列版

            cell.SetCellValue(value);
        }

        //セル設定(数値用)
        public static void WriteCell(ISheet sheet, int columnIndex, int rowIndex, double value)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellValue(value);
        }
    }


}

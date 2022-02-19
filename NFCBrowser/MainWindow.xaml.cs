using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data;
using PCSC;
using PCSC.Iso7816;
using System.Data.SQLite;
using System.IO;
using System.Reflection;

namespace NFCBrowser
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        DataTable cardHistory = new DataTable();
        String cardReaderName = "";
        String UID = "";
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// カード読み込み
        /// </summary>
        private void ReadCard_Click(object sender, RoutedEventArgs e)
        {
            TextBox_Log.Clear();
            TextBox_User.Clear();
            ListBox_DateList.Items.Clear();

            var cardReader = new NFCReader();


            try
            {

                // TextBox設定
                StringBuilder textSB = new StringBuilder();

                cardReaderName = cardReader.GetFirstReaderName();
                textSB.Append(@"Reader names: { " + cardReaderName + " }\r\n");
                textSB.Append("-----------------------------------------------\r\n");

                UID = cardReader.GetUID();
                textSB.Append(@"UID: { " + UID + " }\r\n");
                textSB.Append("-----------------------------------------------\r\n");
                SQLiteDB nfcdb = new SQLiteDB("NFCBrowser.db");
                TextBox_User.Text = nfcdb.selectUser(UID);

                cardHistory = cardReader.GetCardHistory();

                // 支払金額設定
                int expense = 0;
                for (int i = 0; i < cardHistory.Rows.Count - 1; i++)
                {
                    expense = int.Parse(cardHistory.Rows[i + 1]["zandaka"].ToString()) - int.Parse(cardHistory.Rows[i]["zandaka"].ToString());
                    cardHistory.Rows[i]["Shiharai"] = expense;
                }
                // cardHistoryの最後の行[Shiharai]に0を入れる
                cardHistory.Rows[cardHistory.Rows.Count - 1]["Shiharai"] = 0;

                // ListBox設定
                String tmpDate = "";
                for (int i = 0; i < cardHistory.Rows.Count; i++)
                {
                    if (tmpDate != (string)cardHistory.Rows[i]["Date"])
                    {
                        ListBox_DateList.Items.Add(cardHistory.Rows[i]["Date"]);
                    }
                    tmpDate = (string)cardHistory.Rows[i]["Date"];
                }


                for (int i = 0; i < cardHistory.Rows.Count; i++)
                {

                    
                    textSB.Append(@"ROW Data: { " + cardHistory.Rows[i]["Nama"] + " }\r\n");
                    textSB.Append(@"日付: { " + cardHistory.Rows[i]["Date"] + " }\t\t");
                    textSB.Append(@"端末: { " + cardHistory.Rows[i]["Tanmatu"] + " }\t");
                    textSB.Append(@"処理: { " + cardHistory.Rows[i]["Proc"] + " }\r\n");
                    textSB.Append(@"入: { " + cardHistory.Rows[i]["IriSen"] + " ");
                    textSB.Append(@":  " + cardHistory.Rows[i]["Iri"] + " }\t");
                    textSB.Append(@"出: { " + cardHistory.Rows[i]["DeSen"] + " ");
                    textSB.Append(@":  " + cardHistory.Rows[i]["De"] + " }　");
                    textSB.Append(@"支払: { " + cardHistory.Rows[i]["Shiharai"] + " }　");
                    textSB.Append(@"残高: { " + cardHistory.Rows[i]["Zandaka"] + " }\r\n");
                    // textSB.Append(@"連番: { " + cardHistory.Rows[i]["Renban"] + " }\r\n");
                    textSB.Append("-----------------------------------------------\r\n");
                }
                TextBox_Log.Text += textSB.ToString();

            }
            catch (Exception ex)
            {
                TextBox_Log.Text = @"例外が発生： " + ex.Message + "\r\n";
                
                return;
            }

            TextBox_Log.Text += "-----------------------------------------------\r\n";
        }
        /// <summary>
        /// エクセルファイルに書き込む
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PrintCard_Click(object sender, RoutedEventArgs e)
        {
            ExcelControl excel = new ExcelControl();

            // 訪問先の入力チェック
            if (TextBox_Cust.Text == "")
            {
                MessageBox.Show("訪問先は必須入力項目");
                return;
            }

            // cardHistoryに選んだものだけfilterかける
            if (ListBox_DateList.SelectedItems.Count == 0)
            {
                MessageBox.Show("日付は、最低一個は選ぶ");
                return;
            }
            List<string> dlist = new List<string>();

            for (int i = 0; i < ListBox_DateList.SelectedItems.Count; i++)
            {
                dlist.Add(ListBox_DateList.SelectedItems[i].ToString());
            }

            // filter 条件作る
            string filterString = "Date = '";
            filterString += dlist[0] + "'";
            for (int i = 1; i < dlist.Count; i++)
            {
                filterString += " OR Date = '";
                filterString += dlist[i];
                filterString += "'";
            }

            DataRow[] dr = cardHistory.Select(filterString, "Renban DESC");

            // ファイルに書き込む
            try
            {
                // templateをtmpに上書きコピー
                File.Copy(@"template.xlsx", @"tmp.xlsx", true);

                excel.OpenWorkbook(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),"tmp.xlsx"));
                excel.SetICOCA(dr);
                Dictionary<string, string> dc = new Dictionary<string, string>();
                dc.Add("name", TextBox_User.Text);
                dc.Add("cust", TextBox_Cust.Text);
                DateTime date = DateTime.Now;
                dc.Add("year", date.Year.ToString());
                dc.Add("month", date.Month.ToString());
                dc.Add("day",date.Day.ToString());
                double total = 0;
                foreach (var item in dr)
                {
                    total += double.Parse(item["Shiharai"].ToString());
                }
                dc.Add("total", total.ToString());
                excel.SetOthers(dc);

                excel.CloseWorkbook();
            }
            catch(Exception ex)
            {
                TextBox_Log.Text += @"例外が発生： " + ex.Message + "\r\n";
                return;
            }

            // エクセルファイル表示
            System.Diagnostics.Process.Start("ICOCA利用明細書.xlsx");
        }


    }
}

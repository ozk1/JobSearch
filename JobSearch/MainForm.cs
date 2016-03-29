using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.Web;
using HtmlAgilityPack;

namespace WindowsFormsApplication1
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        #region // ハローワーク //
        private void button1_Click(object sender, EventArgs e)
        {
            // リンク取得
            string strValue = textBox1.Text.Replace("\r\n", "\n");
            List<string> LinkList = new List<string>(strValue.Split('\n'));

            // Webからデータ取得
            DataTable TempTable = new DataTable();
            foreach (string link in LinkList)
            {
                if (link.Trim() == "")
                {
                    continue;
                }
                var InfoDict = GetInfo(link);

                // カラムが無かったら追加
                foreach (string Col in InfoDict.Keys)
                {
                    if (TempTable.Columns.Contains(Col) == false)
                    {
                        TempTable.Columns.Add(Col, typeof(string));
                    }
                }

                // 1行追加
                var AddRow = TempTable.Rows.Add();
                foreach (string Col in InfoDict.Keys)
                {
                    AddRow[Col] = InfoDict[Col];
                }

                // 少し待つ
                System.Threading.Thread.Sleep(1000);
            }
            ConvertDataTableToCsv(TempTable, @"D:\test1.csv", true);
        }

        private Dictionary<string, string> GetInfo(string Url)
        {
            Dictionary<string, string> RetDict = new Dictionary<string, string>();
            try
            {
                WebClient webclient = new WebClient();
                webclient.Encoding = Encoding.UTF8;
                string htmlText = webclient.DownloadString(Url);

                if (htmlText != null)
                {
                    // HtmlDocumentオブジェクトを構築する
                    var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(htmlText);

                    var nodes = htmlDoc.DocumentNode.SelectNodes("//tr");
                    foreach (HtmlNode node in nodes)
                    {
                        string Title = "";
                        string Value = "";
                        foreach (HtmlNode childnode in node.ChildNodes)
                        {
                            if (childnode.Name == "th")
                            {
                                Title = TrimEx(childnode.InnerText);
                            }
                            else
                            {
                                Value += TrimEx(childnode.InnerText);
                            }
                        }
                        RetDict[Title] = HttpUtility.HtmlDecode(Value);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.Message);
            }
            return RetDict;
        }

        private string TrimEx(string str)
        {
            str = str.Replace("\t", "");
            str = str.Replace("\n", "");
            str = str.Replace("\r", "");
            return str;
        }
        #endregion

        #region // en転職 //
        private void button2_Click(object sender, EventArgs e)
        {
            // リンク取得
            string strValue = textBox1.Text.Replace("\r\n", "\n");
            List<string> LinkList = new List<string>(strValue.Split('\n'));

            // リンクから詳細ページリスト取得
            DataTable TempTable = new DataTable();
            List<string> DetailPageList = new List<string>();
            foreach (string link in LinkList)
            {
                if (link.Trim() == "")
                {
                    continue;
                }
                DetailPageList.AddRange(GetEnList(link));
            }

            // 詳細ページから内容取得
            foreach (string page in DetailPageList)
            {
                if(page.Trim() == "")
                {
                    continue;
                }

                var InfoDict = GetEnDetail(page);

                // カラムが無かったら追加
                foreach (string Col in InfoDict.Keys)
                {
                    if (TempTable.Columns.Contains(Col) == false)
                    {
                        TempTable.Columns.Add(Col, typeof(string));
                    }
                }

                // 1行追加
                var AddRow = TempTable.Rows.Add();
                foreach (string Col in InfoDict.Keys)
                {
                    AddRow[Col] = InfoDict[Col];
                }

                // 少し待つ
                System.Threading.Thread.Sleep(1000);
            }
            ConvertDataTableToCsv(TempTable, @"D:\test2.csv", true);
        }

        private List<string> GetEnList(string Url)
        {
            List<string> RetList = new List<string>();
            try
            {
                WebClient webclient = new WebClient();
                webclient.Encoding = Encoding.UTF8;
                string htmlText = webclient.DownloadString(Url);

                if (htmlText != null)
                {
                    // HtmlDocumentオブジェクトを構築する
                    var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(htmlText);

                    var nodes = htmlDoc.DocumentNode.SelectNodes("//a");
                    foreach (HtmlNode node in nodes)
                    {
                        if (node.InnerText == "詳細へ")
                        {
                            var targets = node.Attributes.Where(n => n.Name == "href");
                            if(targets != null)
                            {
                                foreach (HtmlAttribute attr in targets)
                                {
                                    RetList.Add(string.Format("https://employment.en-japan.com{0}", attr.Value));
                                }                                
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.Message);
            }
            return RetList;
        }

        private Dictionary<string,string> GetEnDetail(string Url)
        {
            List<string> TargetItems = new List<string>() {"応募資格", "休日休暇", "給与", "勤務時間", "勤務地・交通",
                "雇用形態", "仕事内容", "事業内容", "従業員数", "職種名", "代表者", "入社までの流れ", "福利厚生",
                "募集背景", "連絡先", "ホームページ", "設立", "応募受付方法", "資本金", "事業所", "面接地",
                "売上高", "主要取引先", "教育制度", "関連会社", "配属部署" };
            
            Dictionary<string, string> RetDict = new Dictionary<string, string>();
            try
            {
                WebClient webclient = new WebClient();
                webclient.Encoding = Encoding.UTF8;
                string htmlText = webclient.DownloadString(Url);

                if (htmlText != null)
                {
                    // HtmlDocumentオブジェクトを構築する
                    var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.LoadHtml(htmlText);

                    var nodes = htmlDoc.DocumentNode.SelectNodes("//tr");
                    foreach (HtmlNode node in nodes)
                    {
                        string Title = "";
                        string Value = "";
                        foreach (HtmlNode childnode in node.ChildNodes)
                        {
                            if (childnode.Name == "th")
                            {
                                Title = TrimEx(childnode.InnerText);
                                System.Diagnostics.Trace.WriteLine(Title);
                            }
                            else
                            {
                                Value += TrimEx(childnode.InnerText);
                            }
                        }
                        if (Title != "" && TargetItems.Contains(Title))
                        {
                            RetDict[Title] = HttpUtility.HtmlDecode(Value);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.Message);
            }
            return RetDict;
        }

        #endregion

        #region // DataTable2CSV //
        /// <summary>
        /// DataTableの内容をCSVファイルに保存する
        /// </summary>
        /// <param name="dt">CSVに変換するDataTable</param>
        /// <param name="csvPath">保存先のCSVファイルのパス</param>
        /// <param name="writeHeader">ヘッダを書き込む時はtrue。</param>
        public void ConvertDataTableToCsv(DataTable dt, string csvPath, bool writeHeader)
        {
            //CSVファイルに書き込むときに使うEncoding
            System.Text.Encoding enc =
                System.Text.Encoding.GetEncoding("Shift_JIS");

            //書き込むファイルを開く
            System.IO.StreamWriter sr =
                new System.IO.StreamWriter(csvPath, false, enc);

            int colCount = dt.Columns.Count;
            int lastColIndex = colCount - 1;

            //ヘッダを書き込む
            if (writeHeader)
            {
                for (int i = 0; i < colCount; i++)
                {
                    //ヘッダの取得
                    string field = dt.Columns[i].Caption;
                    //"で囲む
                    field = EncloseDoubleQuotesIfNeed(field);
                    //フィールドを書き込む
                    sr.Write(field);
                    //カンマを書き込む
                    if (lastColIndex > i)
                    {
                        sr.Write(',');
                    }
                }
                //改行する
                sr.Write("\r\n");
            }

            //レコードを書き込む
            foreach (DataRow row in dt.Rows)
            {
                for (int i = 0; i < colCount; i++)
                {
                    //フィールドの取得
                    string field = row[i].ToString();
                    //"で囲む
                    field = EncloseDoubleQuotesIfNeed(field);
                    //フィールドを書き込む
                    sr.Write(field);
                    //カンマを書き込む
                    if (lastColIndex > i)
                    {
                        sr.Write(',');
                    }
                }
                //改行する
                sr.Write("\r\n");
            }

            //閉じる
            sr.Close();
        }

        /// <summary>
        /// 必要ならば、文字列をダブルクォートで囲む
        /// </summary>
        private string EncloseDoubleQuotesIfNeed(string field)
        {
            if (NeedEncloseDoubleQuotes(field))
            {
                return EncloseDoubleQuotes(field);
            }
            return field;
        }

        /// <summary>
        /// 文字列をダブルクォートで囲む
        /// </summary>
        private string EncloseDoubleQuotes(string field)
        {
            if (field.IndexOf('"') > -1)
            {
                //"を""とする
                field = field.Replace("\"", "\"\"");
            }
            return "\"" + field + "\"";
        }

        /// <summary>
        /// 文字列をダブルクォートで囲む必要があるか調べる
        /// </summary>
        private bool NeedEncloseDoubleQuotes(string field)
        {
            return field.IndexOf('"') > -1 ||
                field.IndexOf(',') > -1 ||
                field.IndexOf('\r') > -1 ||
                field.IndexOf('\n') > -1 ||
                field.StartsWith(" ") ||
                field.StartsWith("\t") ||
                field.EndsWith(" ") ||
                field.EndsWith("\t");
        }
        #endregion

        private void button3_Click(object sender, EventArgs e)
        {

        }
    }
}

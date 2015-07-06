using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace GomajiOrderToFridayOrder
{
    public partial class Form1 : Form
    {
        private string appPath_;
        private static Excel.Application _Excel = null;

        public Form1()
        {
            InitializeComponent();

            this.appPath_ = Directory.GetCurrentDirectory();
            textBox1.Text = this.appPath_ + "\\廠商設定.xlsx";
        }

        private void initailExcel()
        {
            //檢查PC有無Excel在執行
            bool flag = false;
            foreach (var item in Process.GetProcesses())
            {
                if (item.ProcessName == "EXCEL")
                {
                    flag = true;
                    break;
                }
            }

            if (!flag)
            {
                _Excel = new Excel.Application();
            }
            else
            {
                object obj = Marshal.GetActiveObject("Excel.Application");//引用已在執行的Excel
                _Excel = obj as Excel.Application;
            }

            _Excel.Visible = true;//設false效能會比較好
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = this.appPath_;
            openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|Excel files 2003~2007 (*.xls)|*.xls|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        using (myStream)
                        {
                            // Insert code to read the stream here.
                            FileStream fs = myStream as FileStream;
                            if (fs != null)
                            {
                                textBox1.Text = fs.Name.ToString();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = this.appPath_;
            openFileDialog1.Filter = "Excel files 2003~2007 (*.xls)|*.xls|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        using (myStream)
                        {
                            // Insert code to read the stream here.
                            FileStream fs = myStream as FileStream;
                            if (fs != null)
                            {
                                textBox2.Text = fs.Name.ToString();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {            
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = this.appPath_;
            saveFileDialog1.FileName = "";
            saveFileDialog1.DefaultExt = ".xls";
            saveFileDialog1.Filter = "Excel files 2003~2007 (*.xls)|*.xls|All files (*.*)|*.*";

            saveFileDialog1.ShowDialog();
            this.textBox3.Text = saveFileDialog1.FileName;
    
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string configPath = this.textBox1.Text;
            string inputPath = this.textBox2.Text;
            string outputPath = this.textBox3.Text;

            this.label1.Text = "";
            this.label2.Text = "";
            this.label3.Text = "";

            if (!File.Exists(configPath))
            {
                label1.Text = "設定檔不存在!";
                label1.ForeColor = Color.Red;
                return;
            }

            if (!File.Exists(inputPath))
            {
                label2.Text = "輸入檔案不存在!";
                label2.ForeColor = Color.Red;
                return;
            }

            this.initailExcel();

            Dictionary<string, string> PIDVendorMap = new Dictionary<string, string>();
            Dictionary<string, Dictionary<string, string>> VendorCfg = new Dictionary<string, Dictionary<string, string>>();
            if (!this.ReadPICVendorCfg(configPath, ref PIDVendorMap, ref VendorCfg))
            {
                label1.Text = "設定檔讀檔失敗";
                return;
            }

            List<Dictionary<string, string>> OrderList = new List<Dictionary<string, string>>();
            if (!this.ReadGomajiOrder(inputPath, ref OrderList))
            {
                label2.Text = "Gomoji 訂單讀取失敗";
                return;
            }

            if (!this.WriteSudaEOD(outputPath, OrderList, PIDVendorMap, VendorCfg))
            {
                return;
            }

            this.HandlePackageList(OrderList);
        }

        private bool ReadPICVendorCfg(string path, ref Dictionary<string, string> PIDVendorMap, ref Dictionary<string, Dictionary<string, string>> VendorCfg)
        {
            Excel.Workbook book = null;
            Excel.Range range = null;

            try
            {
                book = _Excel.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//開啟舊檔案
                Excel.Sheets excelSheets = _Excel.Worksheets;
                string currentSheet = "PID廠商對應";
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
             
                range = excelWorksheet.UsedRange;
                int lastUsedRow = range.Row + range.Rows.Count;

                for (int r = 1; r < lastUsedRow; ++r)
                {
                    var ProductId = (excelWorksheet.Cells[r, 1] as Excel.Range).Value.ToString();

                    if (PIDVendorMap.ContainsKey(ProductId))
                    {
                        continue;
                    }
                    else
                    {
                        PIDVendorMap.Add(ProductId, (excelWorksheet.Cells[r, 2] as Excel.Range).Value.ToString());
                    }
                }


                currentSheet = "廠商資料";
                excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);

                range = excelWorksheet.UsedRange;
                lastUsedRow = range.Row + range.Rows.Count;

                for (int r = 1; r < lastUsedRow; ++r)
                {
                    var VendorName = (excelWorksheet.Cells[r, 1] as Excel.Range).Value.ToString();

                    if (VendorCfg.ContainsKey(VendorName))
                    {
                        continue;
                    }
                    else
                    {
                        //PIDVendorMap.Add(ProductId, (excelWorksheet.Cells[r, 2] as Excel.Range).Value.ToString());
                        Dictionary<string, string> CfgDetail = new Dictionary<string, string>();
                        CfgDetail.Add("Address", (excelWorksheet.Cells[r, 2] as Excel.Range).Value.ToString());
                        CfgDetail.Add("ContentWindow", (excelWorksheet.Cells[r, 3] as Excel.Range).Value.ToString());
                        CfgDetail.Add("Phone", (excelWorksheet.Cells[r, 4] as Excel.Range).Value.ToString());

                        VendorCfg.Add(VendorName, CfgDetail);
                    }
                }
            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                book.Close(Type.Missing, Type.Missing, Type.Missing);
                book = null;
            }

            return true;
        }

        private bool ReadGomajiOrder(string inpath, ref List<Dictionary<string,string>> OrderList)
        {
            Excel.Workbook book = null;
            Excel.Range range = null;

            try
            {
                book = _Excel.Workbooks.Open(inpath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//開啟舊檔案
                Excel.Sheets excelSheets = _Excel.Worksheets;
                string currentSheet = "Worksheet";
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);

                range = excelWorksheet.UsedRange;
                int lastUsedRow = range.Row + range.Rows.Count;
                int lastUsedCol = range.Column;

                string OrderDate = "";
                string OrderID = "";
                string OrderName = "";
                string ReceveName = "";
                string RecevePhone = "";
                string ReceveAddress = "";
                string OrderNote = "";
                string ProductName = "";
                string PID = "";
                string ProductCount = "";

                int OrderCount = 0;

                for (int r = 1; r < lastUsedRow; ++r)
                {
                    if (r == 1)
                    {
                        continue;
                    }

                    //var store_id = Convert.ToString((excelWorksheet.Cells[r, 1] as Excel.Range).Value);
                    Dictionary<string, string> OrderDetail = new Dictionary<string, string>();
                    var tmp  = Convert.ToString((excelWorksheet.Cells[r, 3] as Excel.Range).Value);
                    if (tmp == null)
                    {

                    }
                    else
                    {
                        OrderDate = tmp;
                        OrderCount++;
                        OrderID = String.Format("{0}-{1:0000}", OrderDate, OrderCount);
                        OrderName = Convert.ToString((excelWorksheet.Cells[r, 4] as Excel.Range).Value);
                        ReceveName = Convert.ToString((excelWorksheet.Cells[r, 5] as Excel.Range).Value);
                        RecevePhone = Convert.ToString((excelWorksheet.Cells[r, 7] as Excel.Range).Value);
                        ReceveAddress = Convert.ToString((excelWorksheet.Cells[r, 10] as Excel.Range).Value);
                        OrderNote = Convert.ToString((excelWorksheet.Cells[r, 15] as Excel.Range).Value);
                        OrderNote += Convert.ToString((excelWorksheet.Cells[r, 22] as Excel.Range).Value);
                    }

                    ProductName = Convert.ToString((excelWorksheet.Cells[r, 11] as Excel.Range).Value);
                    PID = Convert.ToString((excelWorksheet.Cells[r, 13] as Excel.Range).Value);
                    ProductCount = Convert.ToString((excelWorksheet.Cells[r, 14] as Excel.Range).Value);

                    OrderDetail.Add("OrderDate", OrderDate);
                    OrderDetail.Add("OrderID", OrderID);
                    OrderDetail.Add("OrderName", OrderName);
                    OrderDetail.Add("ReceveName", ReceveName);
                    OrderDetail.Add("RecevePhone", RecevePhone);
                    OrderDetail.Add("ReceveAddress", ReceveAddress);
                    OrderDetail.Add("OrderNote", OrderNote);
                    OrderDetail.Add("PID", PID);
                    OrderDetail.Add("ProductName", ProductName);
                    OrderDetail.Add("ProductCount", ProductCount);

                    OrderList.Add(OrderDetail);
                }
            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                book.Close(Type.Missing, Type.Missing, Type.Missing);
                book = null;
            }

            return true;
        }

        private bool WriteSudaEOD(string outpath, List<Dictionary<string, string>> OrderList, Dictionary<string, string> PIDVendorMap, Dictionary<string, Dictionary<string, string>> VendorCfg)
        {
            Excel.Workbook book = null;

            try
            {
                book = _Excel.Workbooks.Open(this.appPath_ + "\\order_example.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//開啟舊檔案
                Excel.Sheets excelSheets = _Excel.Worksheets;
                string currentSheet = "Sheet1";
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);

                int r = 1;
                var tmpOrderId = "";
                foreach (var Orderdetail in OrderList)
                {
                    if (tmpOrderId == Orderdetail["OrderID"])
                    {
                        continue;
                    }

                    r++;
                    tmpOrderId = Orderdetail["OrderID"];
                    excelWorksheet.Cells[r, 2].Value = "3";
                    excelWorksheet.Cells[r, 3].Value = "N";
                    excelWorksheet.Cells[r, 4].Value = Orderdetail["ReceveName"];
                    excelWorksheet.Cells[r, 5].Value = Orderdetail["RecevePhone"];
                    excelWorksheet.Cells[r, 6].Value = Orderdetail["RecevePhone"];
                    excelWorksheet.Cells[r, 7].Value = Orderdetail["ReceveAddress"];
                    if (!PIDVendorMap.ContainsKey(Orderdetail["PID"]))
                    {
                        label3.Text = string.Format("找不到 {0} 對應廠商, 請修改 廠商設定.xlsx 後再次嘗試", Orderdetail["PID"]);
                        return false;
                    }

                    var VendorName = PIDVendorMap[Orderdetail["PID"]];
                    if (!VendorCfg.ContainsKey(VendorName))
                    {
                        label3.Text = string.Format("找不到 {0} 廠商資料, 請修改 廠商設定.xlsx 後再次嘗試", VendorName);
                        return false;
                    }

                    excelWorksheet.Cells[r, 8].Value = VendorCfg[VendorName]["ContentWindow"];
                    excelWorksheet.Cells[r, 9].Value = VendorCfg[VendorName]["Phone"];
                    excelWorksheet.Cells[r, 10].Value = VendorCfg[VendorName]["Phone"];
                    excelWorksheet.Cells[r, 11].Value = VendorCfg[VendorName]["Address"];
                    excelWorksheet.Cells[r, 12].Value = "4";
                    excelWorksheet.Cells[r, 13].Value = Orderdetail["ProductName"];
                    excelWorksheet.Cells[r, 14].Value = Orderdetail["OrderNote"];
                }
                
            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                book.SaveAs(outpath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                book.Close(Type.Missing, Type.Missing, Type.Missing);
                book = null;
            }

            return true;
        }

        private bool HandlePackageList(List<Dictionary<string, string>> OrderList)
        {
            Dictionary<string, int> SaleCount = new Dictionary<string, int>();
            foreach (var OrderDetail in OrderList)
            {
                if (SaleCount.ContainsKey(OrderDetail["PID"]))
                {
                    SaleCount[OrderDetail["PID"]] += Convert.ToInt32(OrderDetail["ProductCount"]);
                }
                else
                {
                    SaleCount.Add(OrderDetail["PID"], Convert.ToInt32(OrderDetail["ProductCount"]));
                }
            }

            Excel.Workbook book = null;
            try
            {
                book = _Excel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet ws = (Excel.Worksheet)book.Worksheets[1];

                int r = 1;
                ws.Cells[r, 1].Value = "商品PID";
                ws.Cells[r, 2].Value = "數量";
                foreach (var SaleInfo in SaleCount)
                {
                    r++;
                    ws.Cells[r, 1].Value = SaleInfo.Key;
                    ws.Cells[r, 2].Value = SaleInfo.Value;
                }
            }
            catch (System.Exception ex)
            {
                return false;
            }
            finally
            {
                string outpath = string.Format("{0}\\撿貨清單\\撿貨清單{1}.xlsx", this.appPath_, DateTime.Now.ToString("yyyy-MM-dd"));

                //book.Save();
                book.SaveAs(outpath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                book.Close(Type.Missing, Type.Missing, Type.Missing);
                book = null;
            }
            
            return true;
        }
    }
}

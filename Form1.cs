using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.IO;

using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;
using System.Collections;

namespace sentiment analysis
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
            {
                filePath = file.FileName; //get the path of the file  
                fileExt = Path.GetExtension(filePath); //get the file extension  
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        System.Data.DataTable dtExcel = new System.Data.DataTable();
                        dtExcel.Columns.Add("ID"); dtExcel.Columns.Add("Opinion");
                        dtExcel = ReadExcel(filePath, fileExt); //read excel file 

                        dataGridView1.Visible = true;

                        dataGridView1.DataSource = dtExcel;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                }
            }
        }

        public System.Data.DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            System.Data.DataTable dtexcel = new System.Data.DataTable();

            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch { }
            }
            return dtexcel;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int z = dataGridView1.RowCount;
            // string[] splitword = null;
            string[] authors = null;
            TextWriter tw = new StreamWriter(@"word score.txt");
          

            ///////////////////////////////////
            MessageBox.Show(z.ToString());
            ArrayList listw = new ArrayList();
            ArrayList listwc = new ArrayList();
            ArrayList listwcp = new ArrayList();

            char[] commaSeparator = new char[] { ':', '!', '.', '?' };
            int row = 1;
           // int zero = 0;

            //سطرهای گرید که شامل نظرات هستند را می خواند
            for (int i = 0; i < dataGridView1.RowCount - 2; i++)
                //for (int i = 0; i < 3; i++)
            {
                //جملات یک نظر داخل authoors ریخته می شود
                authors = (dataGridView1[1, i + 1].Value.ToString().Split(commaSeparator));
                ArrayList authorslist = new ArrayList(authors);

                for (int w = 0; w < authorslist.Count; w++)
                {
                    authorslist.Remove(" ");
                    authorslist.Remove("");
                    authorslist.Remove("  ");
                    authorslist.Remove("    ");
                }
                string[] authorss = (string[])authorslist.ToArray(typeof(string));

                for (int j = 0; j < authorss.Length; j++)
                {

                    ArrayList splitword = new ArrayList(authorss[j].Split());
                    for (int k = 0; k < splitword.Count; k++)
                    {
                        splitword.Remove("");
                        splitword.Remove(" ");
                        splitword.Remove("  ");
                        splitword.Remove("   ");
                    }

                    int com = 2;
                    int comn = com - 1;

                    //جملات به کلمات شکسته می شود 
                    for (int k = 0; k < splitword.Count; k++)
                    {
                       // sheet1.Cells[row, com].value2 = splitword[k];
                        listw.Add(splitword[k]);
                        //کلمه داخل اکسل ریخته می شود

                        //گرید را پیمایش می کند تا نمره ی کلمه را بدست آورد
                        for (int q = 0; q < dataGridView2.RowCount - 1; q++)
                        {
                            string s = "";
                            //داخل گرید اگر کلمه ای با * شروع شده بود * را حذف می کند و سپس کلمه ی بدست آمده را داخل متغیر s میریزد
                            if (dataGridView2[0, q].Value.ToString().Substring(dataGridView2[0, q].Value.ToString().Length - 1, 1) == "*")
                            {
                                s = dataGridView2[0, q].Value.ToString().Replace("*", "");
                                //اگر متغیر اس  با کلمه ی داخل اسپلیت ورد یکسان بود از ستون دوم گرید نمره اش را بدست آورده و داخل اکسل ذخیره می کند.(نمره را در سطر زیرین کلمه و در همان ستون باید ذخیره کند)
                                if (splitword[k].ToString() == s)
                                {

                                  //  sheet1.Cells[row + 1, com].value2 = dataGridView2[1, q].Value;
                                    listwc.Add(dataGridView2[1, q].Value);
                                    break;
                                }
                            }
                            // string s = dataGridView2[0, q].Value.ToString().Substring(0,1);
                            // اگر تمامی سطرهای گرید را پیمایش کرد و کلمه را پیدا نکرد نمره ی آن کلمه را صفر قرار بده.البته این قسمت را اجرا نمی کند
                            if (q == dataGridView2.RowCount - 2 && dataGridView2[0, q].Value != splitword[k])
                            {
                              //  sheet1.Cells[row + 1, com].value2 = zero.ToString();
                                listwc.Add(0);
                                break;
                            }
                            if (splitword[k].ToString() == (dataGridView2[0, q].Value.ToString()))
                            {
                               // sheet1.Cells[row + 1, com].value2 = dataGridView2[1, q].Value;
                                listwc.Add(dataGridView2[1, q].Value);
                                break;
                            }

                        }//q
                        // تا وقتی جمله تمام نشده کلمات باید داخل یک سطر از اکسل قرار گیرد به همین خاطر فقط متغیر com را اضافه می کنیم       
                        com++;

                    }//k
                    //وقتی جمله تمام شد و جمله ی بعدی مورد برسی قرار گرفت باید متغیر  رو را 2 تا اضافه کنیم.در واقع سطر اول کلمات و سطر دوم نمره ها باز جمله ی دوم در سطر سوم قرار می گیرد
                    row = row + 2;
                    //جمله به جمله اطلاعات را ذخیره می کند
                    listw.Insert(0, i+1);
                    listwc.Insert(0, i+1);
                    foreach (object item in listw)
                    {
                        tw.Write(item.ToString());
                        tw.Write(" ");
                    }
                    tw.WriteLine("");
                    foreach (object item in listwc)
                    {
                        tw.Write(item.ToString());
                        tw.Write(" ");
                    }
                    tw.WriteLine("");
                    listwc.Clear();
                    listw.Clear();
                    //book1.Save();
                    ///////////////////////////////////////////////////////////////////////        

                }//j
                authorslist.Clear();

            }//i
            tw.Close();
            MessageBox.Show(" ریویو تمام شد");

        }

        private void button3_Click(object sender, EventArgs e)
        {
            ArrayList alist = new ArrayList();
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
            {
                filePath = file.FileName; //get the path of the file  
                fileExt = Path.GetExtension(filePath); //get the file extension  
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        System.Data.DataTable dtExcel = new System.Data.DataTable();
                        // dtExcel.Columns.Add("ID"); dtExcel.Columns.Add("Opinion");
                        dtExcel = ReadExce(filePath, fileExt); //read excel file 

                        dataGridView3.Visible = true;

                        dataGridView3.DataSource = dtExcel;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                }
            }
           // Microsoft.Office.Interop.Excel.Application XL = new Microsoft.Office.Interop.Excel.Application();
           // Workbook book1 = XL.Workbooks.Open(@"sentence score3");
           // Worksheet sheet1 = book1.Sheets[1];
            TextWriter tw = new StreamWriter(@"sentence score.txt");
            int t = 1;
            int zero = 0;
            int max = 0;
            int min = 0;
            int sum = 0;
            ArrayList listw = new ArrayList();
            double rate;


            for (int i = 0; i < (dataGridView3.RowCount - 2); i = i + 2)
            {

                for (int j = 0; j < dataGridView3.ColumnCount - 1; j++)
                {
                    // ArrayList alist = new ArrayList();

                    string s = "";
                    if (dataGridView3[j + 1, i + 1].Value.ToString() == s.ToString())
                    {
                        break;

                    }
                    alist.Add(dataGridView3[j + 1, i + 1].Value.ToString());

                }
                for (int w = 0; w < alist.Count; w++)
                {
                    alist.Remove(zero.ToString());
                }


                //   int sum = 0;
                sum = 0;
                for (int q = 0; q < alist.Count; q++)
                {


                   // sheet1.Cells[t, q + 4].value2 = alist[q];
                    listw.Add(alist[q]);
                    sum = (Convert.ToInt32(alist[q])) + sum;

                }
                // MessageBox.Show(sum.ToString());
                if (sum > 4)
                    sum = 4;
                if (sum < -4)
                    sum = -4;
               // sheet1.Cells[t, 2].value2 = sum.ToString();
              //  listw.Add(sum.ToString());
                // alist.Clear();
                /////////////////////////////////////////////////////////////////////
                if (sum > max)
                    max = sum;
                if (sum < min)
                    min = sum;

                rate = (double)(sum + 4) / 8;


                //dataGridView3[0, i].Value.ToString();
                listw.Insert(0, dataGridView3[0, i].Value.ToString());
                listw.Insert(1, sum.ToString());
                listw.Insert(2, rate.ToString());
                foreach (object item in listw)
                {
                    tw.Write(item.ToString());
                    tw.Write(" ");
                }
                tw.WriteLine("");

                alist.Clear();
                listw.Clear();

                t++;
              //  book1.Save();
            }
            MessageBox.Show(max.ToString() + "max");
            MessageBox.Show(min.ToString() + "min");
            tw.Close();
          //  book1.Close(true);
           // XL.Quit();
           // System.Runtime.InteropServices.Marshal.ReleaseComObject(XL);

            MessageBox.Show("تمام");
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
            {
                filePath = file.FileName; //get the path of the file  
                fileExt = Path.GetExtension(filePath); //get the file extension  
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        System.Data.DataTable dtExcel = new System.Data.DataTable();
                        //  dtExcel.Columns.Add("ID"); dtExcel.Columns.Add("Opinion");
                        dtExcel = ReadExce(filePath, fileExt); //read excel file 

                        dataGridView2.Visible = true;

                        dataGridView2.DataSource = dtExcel;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                }
            }
        }

        public System.Data.DataTable ReadExce(string fileName, string fileExt)
        {
            string conn = string.Empty;
            System.Data.DataTable dtexcel = new System.Data.DataTable();

            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                    MessageBox.Show("فایل باز شده است");
                }
                catch { }
            }
            return dtexcel;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //  ArrayList alist = new ArrayList();
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
            {
                filePath = file.FileName; //get the path of the file  
                fileExt = Path.GetExtension(filePath); //get the file extension  
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        System.Data.DataTable dtExcel = new System.Data.DataTable();
                        // dtExcel.Columns.Add("ID"); dtExcel.Columns.Add("Opinion");
                        dtExcel = ReadExce(filePath, fileExt); //read excel file 

                        dataGridView4.Visible = true;

                        dataGridView4.DataSource = dtExcel;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                }
            }
            int h = dataGridView4.RowCount;
            MessageBox.Show(h.ToString());
            //////////////////////////////////////////////////////////////////////////
           // Microsoft.Office.Interop.Excel.Application XL = new Microsoft.Office.Interop.Excel.Application();
            //Workbook book1 = XL.Workbooks.Open(@"C:\code\original dempster4");
            //Worksheet sheet1 = book1.Sheets[1];
            TextWriter tw = new StreamWriter(@"original dempster.txt");
            int tt = 1;
            int f = 0;
            int k = 1;
            ArrayList listw = new ArrayList();
            //حلقه زیر به تعداد جملات نظرات، نمره نظر تولید شده برای هر جمله که عددی بین 0 و 1 است را درون ستون اول از شیت4 قرار می دهد.
            int s = 1;
            int star = 0;
            double r;
            double rr;
            int rstar = 0;
            int maxstar = 1;
            int minstar = 5;
            int tn = 0;
            /// حلقه زیر مربوط به کد دمپستر است
            while (s < dataGridView4.RowCount)
            {
               // listw.Add(dataGridView4[2, s-1].Value.ToString());
                // حلقه زیر می گوید تا زمانی که شماره ستون اول از گرید 4 با حرف کا که شماره نظر است برابر بود. یعنی میخواهد نظرات را از هم تشخیص دهد. 
                while (Convert.ToInt32(dataGridView4[0, s - 1].Value) == k)
                {
                    //برای قرار دادن شماره نظر مربوطه در ستون اول از شیت4
                   // sheet1.Cells[s, 1].value2 = k.ToString();
                   // listw.Insert(0, k.ToString());
                    // اگر اف مساوی 0 باشد یعنی جمله اول از نظر موجود
                    if (f == 0)
                    {
                        ////////////////////////////////////////////////////////ایف زیر برای وقتی است که نمره جمله اول از یک نظر 0 بدست آمده و باعث می شود نمره تجمیع که همان ام است 0 شود. پس 0.001 میگذاریمش
                        if ((double)dataGridView4[2, s - 1].Value == 0.0)
                            dataGridView4[2, s - 1].Value = 0.001;
                        ////////////////////////////////////////////////////برای جمله اول از یک نظر نمره جمله اول برابر نمره تجمیع است
                      //  sheet1.Cells[s, 4].value2 = dataGridView4[2, s - 1].Value.ToString();
                        listw.Add(dataGridView4[2, s - 1].Value.ToString());

                        publicm.m = (double)dataGridView4[2, s - 1].Value;
                    }
                    // یعنی جملات دوم تا آخر از نظر موجود
                    if (f != 0)
                    {
                        //////////////////////////////////////////////////////////////
                        if ((double)dataGridView4[2, s - 1].Value == 0.0)//مثل ایف بالا فقط برای جملات دوم تا آخر
                            dataGridView4[2, s - 1].Value = 0.001;

                        /////////////////////////////////////////////////////////////ایف زیر فرمول اصلی دمپستر است که استاد داده
                        if (((1 - publicm.m) * (double)dataGridView4[2, s - 1].Value + publicm.m * (1 - (double)dataGridView4[2, s - 1].Value)) != 1)
                        {
                            publicm.m = publicm.m * (double)dataGridView4[2, s - 1].Value / (1 - ((1 - publicm.m) * (double)dataGridView4[2, s - 1].Value + publicm.m * (1 - (double)dataGridView4[2, s - 1].Value)));
                            listw.Add(publicm.m);
                        }
                        //else
                        //{
                        //    publicm.m = 0;
                        //    sheet1.Cells[s, 2].value2 = publicm.m;
                        //}
                    }
                   // book1.Save();
                    /////////////////////////////////////////////////////////////////////////////برای بدست آوردن ستاره یک نمره از نظر
                    rr = (double)dataGridView4[2, s - 1].Value * 4;
                    rstar = (int)Math.Round(rr + 1);
                    //sheet1.Cells[s, 3] = rstar.ToString();
                    listw.Add(rstar.ToString());
                    foreach (object item in listw)
                    {
                        tw.Write(item.ToString());
                        tw.Write(" ");
                    }
                    // book1.Save();
                    tw.WriteLine("");
                    listw.Clear();
                    ////////////////////////////////////////////////////////برای بدست آوردن تناقض
                    if (Convert.ToInt32(rstar) > maxstar)
                        maxstar = Convert.ToInt32(rstar);
                    if (Convert.ToInt32(rstar) < minstar)
                        minstar = Convert.ToInt32(rstar);
                    /////////////////////////////////////////////////////////////////////////////
                    ////////////////////////////////////////////////////////برای بردن نمره تجمیع در فرم 5 ستاره
                    r = publicm.m * 4;
                    star = (int)Math.Round(r + 1);
                    //sheet1.Cells[s, 5] = star.ToString();

                    f++;
                    s++;
                }
                if (maxstar - minstar > 3)
                    tn++;
                //MessageBox.Show(tn.ToString() + "   " + k.ToString());
                maxstar = 1;
                minstar = 5;
                ///// حرف کا نشان دهنده تغییر شماره نظر است
                k++;
                if (s != 1)
                {
                   // sheet1.Cells[s-1, 5] = star.ToString();
                    listw.Add(star.ToString());
                
                }
 
                // حرف اف نشان دهنده جمله اول بودن از نظر جدید است
                f = 0;
            }
            MessageBox.Show(tn.ToString());

            MessageBox.Show("تمام");

        }

        private void button6_Click(object sender, EventArgs e)
        {
            int z = dataGridView1.RowCount;
            int dd;
            int d;
            // string[] splitword = null;
            string[] authors = null;
                Microsoft.Office.Interop.Excel.Application XL = new Microsoft.Office.Interop.Excel.Application();
              Workbook book1 = XL.Workbooks.Open(@"New Microsoft Excel Worksheet");
            Worksheet sheet1 = book1.Sheets[1];

            ///////////////////////////////////
            MessageBox.Show(z.ToString());

            char[] commaSeparator = new char[] { ':', '!', '.', '?' };
            int row = 1;
            int zero = 0;

            //سطرهای گرید که شامل نظرات هستند را می خواند
            for (int i = 2401; i < 3000; i++)
            {
                //جملات یک نظر داخل authoors ریخته می شود
                authors = (dataGridView1[1, i + 1].Value.ToString().Split(commaSeparator));
                ArrayList authorslist = new ArrayList(authors);

                for (int w = 0; w < authorslist.Count; w++)
                {
                    authorslist.Remove(" ");
                    authorslist.Remove("");
                    authorslist.Remove("  ");
                }
                string[] authorss = (string[])authorslist.ToArray(typeof(string));

                for (int j = 0; j < authorss.Length; j++)
                {

                    ArrayList splitword = new ArrayList(authorss[j].Split());
                    for (int k = 0; k < splitword.Capacity; k++)
                    {
                        splitword.Remove("");

                    }

                       int com = 2;
                     int comn = com - 1;
                     sheet1.Cells[row, comn].value2 = (i + 1).ToString();
                     sheet1.Cells[row + 1, comn].value2 = (i + 1).ToString();
                     int[] stars;
                     stars = new int[5];
                    //////////////////////////////////////////////////////////////////////////////////////////////////////
                    //جملات به کلمات شکسته می شود 
                    for (int k = 0; k < splitword.Count; k++)
                    {
                        //  sheet1.Cells[row, com].value2 = splitword[k];

                        //کلمه داخل اکسل ریخته می شود

                        //گرید را پیمایش می کند تا نمره ی کلمه را بدست آورد
                        for (int q = 0; q < dataGridView2.RowCount - 1; q++)
                        {
                            string s = "";
                            //داخل گرید اگر کلمه ای با * شروع شده بود * را حذف می کند و سپس کلمه ی بدست آمده را داخل متغیر s میریزد
                            if (dataGridView2[0, q].Value.ToString().Substring(dataGridView2[0, q].Value.ToString().Length - 1, 1) == "*")
                            {
                                s = dataGridView2[0, q].Value.ToString().Replace("*", "");
                                //اگر متغیر اس  با کلمه ی داخل اسپلیت ورد یکسان بود از ستون دوم گرید نمره اش را بدست آورده و داخل اکسل ذخیره می کند.(نمره را در سطر زیرین کلمه و در همان ستون باید ذخیره کند)
                                if (splitword[k].ToString() == s)
                                {

                                    //                        sheet1.Cells[row + 1, com].value2 = dataGridView2[1, q].Value;
                                    //////////////////////////////////////////////////////////////////////////////////////
                                    dd = Convert.ToInt32(dataGridView1[0, i + 1].Value);
                                    d = Convert.ToInt32(dataGridView2[dd + 6, q].Value);
                                    d = d + 1;
                                    dataGridView2[dd + 6, q].Value = d;

                                    //////////////////////////////////////////////////////////////////////////////////////
                                    break;
                                }
                            }
                           //  اگر تمامی سطرهای گرید را پیمایش کرد و کلمه را پیدا نکرد نمره ی آن کلمه را صفر قرار بده.البته این قسمت را اجرا نمی کند
                            if (q == dataGridView2.RowCount - 2 && dataGridView2[0, q].Value != splitword[k])
                            {
                                sheet1.Cells[row + 1, com].value2 = zero.ToString();
                                break;
                            }
                            if (splitword[k].ToString() == (dataGridView2[0, q].Value.ToString()))
                            {
                                //////////////////////////////////////////////////////////////////////////////////////
                                dd = Convert.ToInt32(dataGridView1[0, i + 1].Value);
                                d = Convert.ToInt32(dataGridView2[dd + 6, q].Value);
                                d = d + 1;
                                dataGridView2[dd + 6, q].Value = d;

                                //////////////////////////////////////////////////////////////////////////////////////
                               // sheet1.Cells[row + 1, com].value2 = dataGridView2[1, q].Value;
                                break;
                            }

                        }//q
                        // تا وقتی جمله تمام نشده کلمات باید داخل یک سطر از اکسل قرار گیرد به همین خاطر فقط متغیر com را اضافه می کنیم       
                        com++;

                    }//k
                    //وقتی جمله تمام شد و جمله ی بعدی مورد برسی قرار گرفت باید متغیر  رو را 2 تا اضافه کنیم.در واقع سطر اول کلمات و سطر دوم نمره ها باز جمله ی دوم در سطر سوم قرار می گیرد
                    row = row + 2;
                    //جمله به جمله اطلاعات را ذخیره می کند
                      book1.Save();
                    ///////////////////////////////////////////////////////////////////////        

                }//j
                authorslist.Clear();

            }//i

             book1.Close(true);
            XL.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(XL);
            MessageBox.Show(" ریویو تمام شد");
        }

        private void button8_Click(object sender, EventArgs e)
        {

            TextWriter tw = new StreamWriter(@"ID sentence score.txt");
            ArrayList listw = new ArrayList();
            double sum;
            //double zero = 0.0;
            // int jj ;
            //  int n;
            ArrayList list1 = new ArrayList();
            ArrayList list2 = new ArrayList();
            ArrayList list3 = new ArrayList();
            ArrayList list4 = new ArrayList();
            ArrayList list5 = new ArrayList();
            MessageBox.Show(dataGridView1[1, 0].Value.ToString());

            for (int i = 0; i < (dataGridView1.RowCount - 1); i++)
            {
                //n = 0;
                for (int j = 0; j < dataGridView1.ColumnCount - 1; j++)
                {
                    if (dataGridView1[j + 1, i].Value.ToString() != "")
                    {
                        //n = 0;
                        for (int q = 0; q < (dataGridView2.RowCount - 1); q++)
                        {

                            string s = "";
                            //داخل گرید اگر کلمه ای با * شروع شده بود * را حذف می کند و سپس کلمه ی بدست آمده را داخل متغیر s میریزد
                            if (dataGridView2[0, q].Value.ToString().Substring(dataGridView2[0, q].Value.ToString().Length - 1, 1) == "*")
                            {
                                s = dataGridView2[0, q].Value.ToString().Replace("*", "");
                                //اگر متغیر اس  با کلمه ی داخل اسپلیت ورد یکسان بود از ستون دوم گرید نمره اش را بدست آورده و داخل اکسل ذخیره می کند.(نمره را در سطر زیرین کلمه و در همان ستون باید ذخیره کند)
                                if (dataGridView1[j + 1, i].Value.ToString() == s)
                                {
                                    list1.Add((double)dataGridView2[13, q].Value);
                                    list2.Add((double)dataGridView2[14, q].Value);
                                    list3.Add((double)dataGridView2[15, q].Value);
                                    list4.Add((double)dataGridView2[16, q].Value);
                                    list5.Add((double)dataGridView2[17, q].Value);

                                    //          n = 1;

                                    // sheet1.Cells[row + 1, com].value2 = zero.ToString();
                                    break;
                                }
                            }
                            // string s = dataGridView2[0, q].Value.ToString().Substring(0,1);
                            // اگر تمامی سطرهای گرید را پیمایش کرد و کلمه را پیدا نکرد نمره ی آن کلمه را صفر قرار بده.البته این قسمت را اجرا نمی کند
                           
                            if (dataGridView1[j + 1, i].Value.ToString() == (dataGridView2[0, q].Value.ToString()))
                            {

                                list1.Add((double)dataGridView2[13, q].Value);
                                list2.Add((double)dataGridView2[14, q].Value);
                                list3.Add((double)dataGridView2[15, q].Value);
                                list4.Add((double)dataGridView2[16, q].Value);
                                list5.Add((double)dataGridView2[17, q].Value);
                                //    n = 1;
                                // sheet1.Cells[row + 1, com].value2 = dataGridView2[1, q].Value;
                                break;
                            }

                        }//q
                    }
                }//j

                ////////////////////////////////////////////////////////
               // sheet1.Cells[i + 1, 1].value2 = dataGridView1[0, i].Value;
                listw.Add(dataGridView1[0, i].Value);
                sum = 0.0;
                int c = 0;
                if (list1.Count == 0)
                {
                  //  sheet1.Cells[i + 1, 2].value2 = sum;
                    listw.Add(sum);
                }
                else
                {
                    for (int q = 0; q < list1.Count; q++)
                    {

                        sum = ((double)(list1[q])) + sum;

                    }

                    c = list1.Count;
                   // sheet1.Cells[i + 1, 2].value2 = sum / (double)c;
                    listw.Add(sum / (double)c);
                    list1.Clear();
                }
              //  book1.Save();
                ///////////////////////////////////////////////////
                sum = 0.0;
                if (list2.Count == 0)
                {
                    //sheet1.Cells[i + 1, 3].value2 = sum;
                    listw.Add(sum);
                }
                else
                {
                    for (int q = 0; q < list2.Count; q++)
                    {
                        // sheet1.Cells[t, q + 4].value2 = list1[q];
                        sum = ((double)(list2[q])) + sum;

                    }

                    //sheet1.Cells[i + 1, 1].value2 = i + 1;
                    c = list2.Count;
                   // sheet1.Cells[i + 1, 3].value2 = sum / (double)c;
                    listw.Add(sum / (double)c);
                    list2.Clear();
                }
              //  book1.Save();
                ////////////////////////////////////////////////////
                sum = 0.0;
                if (list3.Count == 0)
                {
                   // sheet1.Cells[i + 1, 4].value2 = sum;
                    listw.Add(sum);
                }
                else
                {
                    for (int q = 0; q < list3.Count; q++)
                    {
                        // sheet1.Cells[t, q + 4].value2 = list1[q];
                        sum = ((double)(list3[q])) + sum;

                    }

                    // sheet1.Cells[i + 1, 1].value2 = i + 1;
                    c = list3.Count;
                   // sheet1.Cells[i + 1, 4].value2 = sum / (double)c;
                    listw.Add(sum / (double)c);
                    list3.Clear();
                }
             //   book1.Save();
                //////////////////////////////////////////////////////
                sum = 0.0;
                if (list4.Count == 0)
                {
                   // sheet1.Cells[i + 1, 5].value2 = sum;
                    listw.Add(sum);
                }
                else
                {
                    for (int q = 0; q < list4.Count; q++)
                    {
                        // sheet1.Cells[t, q + 4].value2 = list1[q];
                        sum = ((double)(list4[q])) + sum;

                    }

                    // sheet1.Cells[i + 1, 1].value2 = i + 1;
                    c = list4.Count;
                  //  sheet1.Cells[i + 1, 5].value2 = sum / (double)c;
                    listw.Add(sum / (double)c);
                    list4.Clear();
                }
              //  book1.Save();
                /////////////////////////////////////////////////////////
                sum = 0.0;
                if (list5.Count == 0)
                {
                    //sheet1.Cells[i + 1, 6].value2 = sum;
                    listw.Add(sum);
                }
                else
                {
                    for (int q = 0; q < list5.Count; q++)
                    {
                        // sheet1.Cells[t, q + 4].value2 = list1[q];
                        sum = ((double)(list5[q])) + sum;

                    }

                    //sheet1.Cells[i + 1, 1].value2 = i + 1;
                    c = list5.Count;
                  //  sheet1.Cells[i + 1, 6].value2 = sum / (double)c;
                    listw.Add(sum / (double)c);
                    list5.Clear();
                }
                //book1.Save();
                ////////////////////////////////////////////////////////
                foreach (object item in listw)
                {
                    tw.Write(item.ToString());
                    tw.Write(" ");
                }
                tw.WriteLine("");
                listw.Clear();
            }
            //book1.Close(true);
            //XL.Quit();
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(XL);
            MessageBox.Show("تمام");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            int z = dataGridView1.RowCount;
            // string[] splitword = null;
            string[] authors = null;
            //Microsoft.Office.Interop.Excel.Application XL = new Microsoft.Office.Interop.Excel.Application();
            //Workbook book1 = XL.Workbooks.Open(@"C:\code\sentence seprator5");
            //Worksheet sheet1 = book1.Sheets[1];
            TextWriter tw = new StreamWriter(@"sentence seprator.txt");
            ArrayList listw = new ArrayList();
            ///////////////////////////////////
            MessageBox.Show(z.ToString());

            char[] commaSeparator = new char[] { ':', '!', '.', '?' };
            int row = 1;
            //int zero = 0;

            //سطرهای گرید که شامل نظرات هستند را می خواند
            for (int i = 0; i < 3000; i++)
            {
                //جملات یک نظر داخل authoors ریخته می شود
                authors = (dataGridView1[1, i + 1].Value.ToString().Split(commaSeparator));
                ArrayList authorslist = new ArrayList(authors);

                for (int w = 0; w < authorslist.Count; w++)
                {
                    authorslist.Remove(" ");
                    authorslist.Remove("");
                    authorslist.Remove("  ");
                }
                string[] authorss = (string[])authorslist.ToArray(typeof(string));

                for (int j = 0; j < authorss.Length; j++)
                {

                    ArrayList splitword = new ArrayList(authorss[j].Split());
                    for (int k = 0; k < splitword.Capacity; k++)
                    {
                        splitword.Remove("");

                    }

                    int com = 2;
                    int comn = com - 1;
                   // sheet1.Cells[row, comn].value2 = (i + 1).ToString();
                    listw.Add((i + 1).ToString());
  



                    //جملات به کلمات شکسته می شود 
                    for (int k = 0; k < splitword.Count; k++)
                    {
                      //  sheet1.Cells[row, com].value2 = splitword[k];
                        listw.Add(splitword[k]);
                        
                    }//k
                    //وقتی جمله تمام شد و جمله ی بعدی مورد برسی قرار گرفت باید متغیر  رو را 2 تا اضافه کنیم.در واقع سطر اول کلمات و سطر دوم نمره ها باز جمله ی دوم در سطر سوم قرار می گیرد
                    foreach (object item in listw)
                    {
                        tw.Write(item.ToString());
                        tw.Write(" ");
                    }
                    tw.WriteLine("");
                    listw.Clear();
                    //جمله به جمله اطلاعات را ذخیره می کند
                  //  book1.Save();
                    ///////////////////////////////////////////////////////////////////////        

                }//j
                authorslist.Clear();

            }//i
            //book1.Close(true);
            //XL.Quit();
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(XL);
           MessageBox.Show(" ریویو تمام شد");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
            {
                filePath = file.FileName; //get the path of the file  
                fileExt = Path.GetExtension(filePath); //get the file extension  
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        System.Data.DataTable dtExcel = new System.Data.DataTable();
                        // dtExcel.Columns.Add("ID"); dtExcel.Columns.Add("Opinion");
                        dtExcel = ReadExce(filePath, fileExt); //read excel file 

                        dataGridView1.Visible = true;

                        dataGridView1.DataSource = dtExcel;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                }
            }
            TextWriter tw = new StreamWriter(@"c:\code\zahra\ID Aggregation Score.txt");
           // int tt = 1;
             int f = 0;
            int k = 2401;
            //حلقه زیر به تعداد جملات نظرات، نمره نظر تولید شده برای هر جمله که عددی بین 0 و 1 است را درون ستون اول از شیت4 قرار می دهد.

            int s = 8322;
            double max1=0;
            double max2=0;
            double max3=0;
            double max4=0;
            int index1 = -1;
            int index2 = -1;
            int index3 = -1;
            int index4 = -1;
            int a = 0;
            int b = 0;
            int d = 0;
            double c1=0;
            double c2=0;
            double[,] ff = { { 0, 0, 0, 0}, { 0, 0, 0, 0}, { 0, 0, 0, 0},{ 0, 0, 0, 0}, { 0, 0, 0, 0} };
            double nf =0.5;
            ArrayList listw = new ArrayList();
            int pred=3;
         //   int zero = 0;
           // sheet1.Cells[s, 2].value2 = zero.ToString();
            /// حلقه زیر مربوط به کد دمپستر است
            while (s < 10560)
            {
                f = 0;
                // حلقه زیر می گوید تا زمانی که شماره ستون اول از گرید 4 با حرف کا که شماره نظر است برابر بود. یعنی میخواهد نظرات را از هم تشخیص دهد. 
                while (Convert.ToInt32(dataGridView1[0, s - 1].Value) == k )
                {
               
                    if (f == 0)
                    {
                        index1 = -1;
                        index2 = -1;
                        max1 = (double)(dataGridView1[7, s - 1].Value);
                        max2 = (double)(dataGridView1[8, s - 1].Value);
                        // MessageBox.Show(max1.ToString());
                        for (int i = 1; i < 6; i++)
                        {
                            if (max1 == (double)(dataGridView1[i, s - 1].Value))
                                if (index1 == -1)
                                    index1 = i - 1;
                                else
                                {
                                    a = index1;
                                    b = i - 1;
                                    d = new Random(DateTime.UtcNow.Millisecond).Next(0, 1);
                                    if (d == 1)
                                        index1 = b;
                                }
                            if (max2 == (double)(dataGridView1[i, s - 1].Value))
                                if (index2 == -1)
                                index2 = i-1;
                                else
                                {
                                    a = index2;
                                    b = i - 1;
                                    d = new Random(DateTime.UtcNow.Millisecond).Next(0, 1);
                                    if (d == 1)
                                        index2 = b;
                                }
                        }
                        c1 = 1 - (double)(dataGridView1[10, s - 1].Value);
                    }
                    if (f != 0)
                    {
                        index3 = -1;
                        index4 = -1;
                        max3 = (double)(dataGridView1[7, s - 1].Value);
                        max4 = (double)(dataGridView1[8, s - 1].Value);
                        for (int i = 1; i < 6; i++)
                        {
                            if (max3 == (double)(dataGridView1[i, s - 1].Value))
                                if (index3 == -1)
                                    index3 = i - 1;
                                else
                                {
                                    a = index3;
                                    b = i - 1;
                                 d=new Random(DateTime.UtcNow.Millisecond).Next(0, 1);
                                 if (d == 1)
                                     index3 = b;
                                }
                            if (max4 == (double)(dataGridView1[i, s - 1].Value))
                                if (index4 == -1)
                                    index4 = i-1;
                                else
                                {
                                    a = index4;
                                    b = i - 1;
                                    d = new Random(DateTime.UtcNow.Millisecond).Next(0, 1);
                                    if (d == 1)
                                        index4 = b;
                                }
                        }
                        c2 = 1 - (double)(dataGridView1[10, s - 1].Value);
                        for (int g = 0; g < 5; g++)
                        {
                            for (int o = 0; o < 4; o++)
                            {
                                ff[g, o] = 0;
                            }
                        }
                        /////////////////////////////////////////////////////////////////////////////////////////////
                        ////////////////////////////////////////////////////////////////////////////////
                        if (index1 != index2 && (index1 == index3 || index1 == index4) && (index2 == index3 || index2 == index4))
                        {
                            if (index1 != index3 && index2 != index4)
                            {
                                nf = 1 - (max1 * max3) - (max2 * max4);
                                ff[index1, 0] = (1 / nf) * (max1 * max4 + max1 * c2 + c1 * max4);
                                ff[index2, 1] = (1 / nf) * (max2 * max3 + max2 * c2 + c1 * max3);
                                c1 = (1 / nf) * c1 * c2;
                            }
                            if (index1 == index3 && index2 == index4)
                            {
                                nf = 1 - (max1 * max4) - (max2 * max3);
                                ff[index1, 0] = (1 / nf) * (max1 * max3 + max1 * c2 + c1 * max3);
                                ff[index2, 1] = (1 / nf) * (max2 * max4 + max2 * c2 + c1 * max4);
                                c1 = (1 / nf) * c1 * c2;
                            }
                        }
                        ////////////////////////////////////////////////////////////////////////////////
                        
                        if (index1 != index2 && index1 != index3 && index2 != index4 && index2 != index3 && index1 != index4 && index3!=index4)
                        {
                            nf = 1 - (max1 * max3) - (max1 * max4) - (max2 * max3) - (max2 * max4);
                            ff[index1, 0] = (1 / nf) * max1 * c2;
                            ff[index2, 1] = (1 / nf) * max2 * c2;
                            ff[index3, 2] = (1 / nf) * max3 * c1;
                            ff[index4, 3] = (1 / nf) * max4 * c1;
                        }
                      ////////////////////////////////////////////////////////////////////////////////////////////////
                        if (index1 != index2 && ((index1 != index3 && index1 == index4 && index2 != index3 && index2 != index4) || (index1 == index3 && index1 != index4 && index2 != index3 && index2 != index4) || (index2 != index3 && index2 == index4 && index1 != index3 && index1 != index4) || (index2 == index3 && index2 != index4 && index1 != index3 && index1 != index4)))
                        {
                            if (index1 == index3)
                            {
                                nf = 1 - (max1 * max4) - (max2*max4)-(max2*max3);
                                ff[index1, 0] = (1 / nf) * (max1 * max3 + max1 * c2 + c1 * max3);
                                ff[index2, 1] = (1 / nf) * max2 * c2;
                                // ff[index3, 2] = (1 / nf) * max3 * c1;
                                ff[index4, 3] = (1 / nf) * max4 * c1;
                                c1 = (1 / nf) * c1 * c2;
                            }
                            if (index1 == index4)
                            {
                                nf = 1 - (max1 * max3) - (max2 * max3) - (max2 * max4);
                                ff[index1, 0] = (1 / nf) * (max1 * max4 + max1 * c2 + c1 * max4);
                                ff[index2, 1] = (1 / nf) * max2 * c2;
                                // ff[index3, 2] = (1 / nf) * max3 * c1;
                                ff[index3, 3] = (1 / nf) * max3 * c1;
                                c1 = (1 / nf) * c1 * c2;
                            }
                            if (index2 == index4)
                            {
                                nf = 1 - (max2 * max3) - (max1 * max3) - (max1 * max4);
                                ff[index2, 0] = (1 / nf) * (max2 * max4 + max2 * c2 + c1 * max4);
                                ff[index1, 1] = (1 / nf) * max1 * c2;
                                // ff[index3, 2] = (1 / nf) * max3 * c1;
                                ff[index3, 3] = (1 / nf) * max3 * c1;
                                c1 = (1 / nf) * c1 * c2;
                            }
                            if (index2 == index3)
                            {
                                nf = 1 - (max2 * max4) - (max1 * max4) - (max1 * max3);
                                ff[index2, 0] = (1 / nf) * (max2 * max3 + max2 * c2 + c1 * max3);
                                ff[index1, 1] = (1 / nf) * max1 * c2;
                                // ff[index3, 2] = (1 / nf) * max3 * c1;
                                ff[index4, 3] = (1 / nf) * max4 * c1;
                                c1 = (1 / nf) * c1 * c2;
                            }
                        }
                        /////////////////////////////////////////////////////////////////////////////////////////////////
                        max1 = ff[0, 0];
                        for (int g = 0; g < 5; g++)
                        {
                            for (int o = 0; o < 4; o++)
                            {
                                if (ff[g,o] > max1)
                                {
                                    max1 = ff[g,o];
                                    index1 = g;
                                }
                            }
                        }
                        max2 = ff[0,0];
                        for (int g = 0; g < 5; g++)
                        {
                            for (int o = 0; o < 4; o++)
                            {
                                if (ff[g,o] > max2 && ff[g,o] != max1)
                                {
                                    max2 = ff[g,o];
                                    index2 = g;
                                }
                            }
                        }
                    c1=1-(max1+max2);
                    }
                   
                     
                    f++;
                    s++;
                  
                }
              
                
                listw.Insert(0, k);
                k++;
                if (s != 1)
                {
                    
                    pred = (index1 + index2) / 2;
                    pred = pred + 1;
                    //sheet1.Cells[s-1, 2].value2 = pred.ToString();
                    listw.Add(pred.ToString());
                 
                    if (index1 > index2)
                        pred = index1;
                    else
                        pred = index2;
                    pred = pred + 1;
                    listw.Add(pred.ToString());
                    listw.Add((index1 + 1).ToString());
                    listw.Add((index2 + 1).ToString());
                }
                foreach (object item in listw)
                {
                    tw.Write(item.ToString());
                    tw.Write(" ");
                }
                tw.WriteLine("");
                listw.Clear();
            }
           
            //book1.Close(true);
            //XL.Quit();
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(XL);

            MessageBox.Show("تمام");
        }
    }
}


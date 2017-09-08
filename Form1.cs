using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;


using System.IO;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using System.Data.OleDb;

namespace Test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

#region 给Excel文档添加页眉页脚

        private List<string> GetExcelSheetNames(string filepath)
        {   
            List<string> tablenames = new List<string>();
            try
            {
                System.Data.DataTable dtTbTmp = new System.Data.DataTable();
                System.Data.DataTable dtSheetNames = new System.Data.DataTable();
                OleDbConnection conn = null;
                String strConnString = String.Empty;               

                //strConnString = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source-={0};Extended Properties=""Excel 12.0;HDR=NO;IMEX=1""", filepath);
                //strConnString = string.Format(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES;IMEX=1'", filepath);
                strConnString = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'", filepath);
                if (!String.IsNullOrEmpty(filepath))
                {
                    conn = new OleDbConnection(strConnString);
                    conn.Open();
                    dtSheetNames = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    conn.Close();
                }

                for (int row = 0; row < dtSheetNames.Rows.Count; row++)
                {
                    string tablename = dtSheetNames.Rows[row]["TABLE_NAME"].ToString().Replace("$", "");
                    tablenames.Add(tablename);
                }

                return tablenames;              
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return tablenames; 
            }
        }

        /// <summary>  
        /// 给Excel文档添加页眉  
        /// </summary>  
        /// <param name="filePath">文件名</param>  
        /// <returns></returns>  
        private void AddPageHeaderFooterForExcel(string filePath)
        {
            OpenFileDialog openFileDlg = new OpenFileDialog();
            openFileDlg.InitialDirectory = "c:";                // 初始目录
            openFileDlg.DefaultExt = "xls";
            openFileDlg.Filter = "Excel文件 (*.xls;*.xlsx)|*.xls*";
            List<string> tnames = new List<string>();
            if (openFileDlg.ShowDialog() == DialogResult.OK)
            {
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbooks wbs = app.Workbooks;
                Microsoft.Office.Interop.Excel.Workbook wb = wbs.Add(openFileDlg.FileName);
                //Microsoft.Office.Interop.Excel.Worksheets wss;
                //Microsoft.Office.Interop.Excel.Worksheet ws;

                //获取Excel中的工作表名
                tnames = GetExcelSheetNames(openFileDlg.FileName);

                //  绘制图片
                int width = 1;
                int height = 1;
                Color backColor = System.Drawing.Color.White;
                Color textColor = System.Drawing.Color.LightCoral;
                String text = "测试文字";
                System.Drawing.Font font = new System.Drawing.Font("宋体", 3);

                Image img = Image.FromFile(@"C:\Users\Administrator\Desktop\1.jpg");
                Bitmap bitmap = new Bitmap(img, new Size(img.Width / 3, img.Height / 3));
                Graphics drawing = Graphics.FromImage(img);
                SizeF textSize = drawing.MeasureString(text, font);
                Brush textBrush = new SolidBrush(textColor);
                drawing.TranslateTransform(((int)width - textSize.Width) / 2, ((int)height - textSize.Height) / 2);
                drawing.RotateTransform(-45);
                drawing.TranslateTransform(-((int)width - textSize.Width) / 2, -((int)height - textSize.Height) / 2);
                drawing.Clear(backColor);
                drawing.DrawString(text, font, textBrush, ((int)width - textSize.Width) / 2, ((int)height - textSize.Height) / 2);
                drawing.Save();

                if (tnames != null)
                {
                    foreach (string tname in tnames)
                    {
                        try
                        {
                            ((Microsoft.Office.Interop.Excel.Worksheet)wb.Sheets[tname.TrimEnd("$".ToCharArray())]).PageSetup.RightHeader = @"&""Arial""&9这是页眉文字";
                            ((Microsoft.Office.Interop.Excel.Worksheet)wb.Sheets[tname.TrimEnd("$".ToCharArray())]).PageSetup.LeftHeaderPicture.Filename = @"C:\Users\Administrator\Desktop\2.jpg";
                            //((Microsoft.Office.Interop.Excel.Worksheet)wb.Sheets[tname.TrimEnd("$".ToCharArray())]).PageSetup.LeftHeaderPicture.Height = 10 ;
                            //((Microsoft.Office.Interop.Excel.Worksheet)wb.Sheets[tname.TrimEnd("$".ToCharArray())]).PageSetup.LeftHeaderPicture.Width = 10;
                            ((Microsoft.Office.Interop.Excel.Worksheet)wb.Sheets[tname.TrimEnd("$".ToCharArray())]).PageSetup.LeftHeader = "&G";

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                }
                app.Save(); 
                app.ActiveWorkbook.Save();
                app.Quit();
                app = null;
                GC.Collect();
            }

        }

        private void AddPageHeaderFooterForExcel2(string filepath)
        {
            //加载Excel文档
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Open(filepath);
            //Microsoft.Office.Interop.Excel.Workbook wb = wbs.Add(@"C:\Users\Administrator\Desktop\Test.xlsx");

            //设置水印文字和字体
            System.Drawing.Font font = new System.Drawing.Font("宋体",10);
            String watermark = "内部资料";

            foreach (Worksheet sheet in wb.Worksheets)
            {
                sheet.PageSetup.RightHeader = watermark;
                sheet.PageSetup.LeftHeaderPicture.Filename = @"C:\Users\Administrator\Desktop\2.jpg";
                sheet.PageSetup.LeftHeader = "&G";
            }
           
            wb.Save();
            wb.Close();
            //app.Save();
            //app.ActiveWorkbook.Save();
            app.Quit();
            app = null;
            GC.Collect();
        }

#endregion 给Excel文档添加页眉页脚  

        private void btExcel_Click(object sender, EventArgs e)
        {
            //AddPageHeaderFooterForExcel(@"C:\Users\Administrator\Desktop\Test.xlsx");
            string filepath = "";
            string filetype = "";
            OpenFileDialog openFileDlg = new OpenFileDialog();
            if (openFileDlg.ShowDialog() == DialogResult.OK)
            {
                filepath = openFileDlg.FileName;
                filetype = Path.GetExtension(filepath);
                if (filetype != ".xls" && filetype != ".xlsx")
                {
                    MessageBox.Show("请选项Excel类型的文件");
                    return;
                }
                tbExcel.Text = filepath;
                AddPageHeaderFooterForExcel2(filepath);
            }
        }

       
#region 给word文档添加页眉页脚
        /// <summary>  
        /// 给word文档添加页眉  
        /// </summary>  
        /// <param name="filePath">文件名</param>  
        /// <returns></returns>  
        public static bool AddPageHeaderFooterForWord(string filePath)
        {
            try
            {
                Object oMissing = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.Application WordApp = new Microsoft.Office.Interop.Word.Application();
                WordApp.Visible = true;
                object filename = filePath;
                Microsoft.Office.Interop.Word._Document WordDoc = WordApp.Documents.Open(ref filename, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                ////添加页眉方法一：  
                //WordApp.ActiveWindow.View.Type = WdViewType.wdOutlineView;  
                //WordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekPrimaryHeader;  
                //WordApp.ActiveWindow.ActivePane.Selection.InsertAfter( "**公司" );//页眉内容  

                ////添加页眉方法二：  
                if (WordApp.ActiveWindow.ActivePane.View.Type == WdViewType.wdNormalView ||
                    WordApp.ActiveWindow.ActivePane.View.Type == WdViewType.wdOutlineView)
                {
                    WordApp.ActiveWindow.ActivePane.View.Type = WdViewType.wdPrintView;
                }
                WordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekCurrentPageHeader;
                WordApp.Selection.HeaderFooter.LinkToPrevious = false;
                //WordApp.Selection.HeaderFooter.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                WordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphDistribute;
                WordApp.Selection.HeaderFooter.Range.Text = "                                             内部公开";
                WordApp.Selection.InlineShapes.AddPicture(@"C:\Users\Administrator\Desktop\2.jpg");

                // 去掉页眉的横线
                //WordApp.ActiveWindow.ActivePane.Selection.ParagraphFormat.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
                //WordApp.ActiveWindow.ActivePane.Selection.Borders[WdBorderType.wdBorderBottom].Visible = false;

                //在页眉的图片后面追加几个字
                //WordApp.ActiveWindow.ActivePane.Selection.InsertAfter("  文档页眉");


                WordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekCurrentPageFooter;
                WordApp.Selection.HeaderFooter.LinkToPrevious = false;
                WordApp.Selection.HeaderFooter.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                WordApp.Selection.HeaderFooter.Range.Text = "页脚内容";
                //WordApp.ActiveWindow.ActivePane.Selection.InsertAfter("页脚内容");

                //跳出页眉页脚设置  
                WordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;

                //保存  
                WordDoc.Save();
                WordDoc.Close(ref oMissing, ref oMissing, ref oMissing);
                WordApp.Quit(ref oMissing, ref oMissing, ref oMissing);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                return false;
            }
        }
#endregion 给word文档添加页眉页脚  


        private void btWord_Click(object sender, System.EventArgs e)
        {
            string filepath = "";
            string filetype = "";
            OpenFileDialog openFileDlg = new OpenFileDialog();
            if (openFileDlg.ShowDialog() == DialogResult.OK)
            {
                filepath = openFileDlg.FileName;
                filetype = Path.GetExtension(filepath);
                if (filetype != ".doc" && filetype != ".docx")
                {
                    MessageBox.Show("请选项Word类型的文件");
                    return;
                }
                tbWord.Text = filepath;
                AddPageHeaderFooterForWord(filepath);

            }
        }


#region 给ppt文档添加页眉页脚

        /// <summary>
        /// 给ppt添加页眉
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool AddPageHeaderFooterForPPT(string filepath)
        {
            try
            {
                //加载PPT文档
                Microsoft.Office.Interop.PowerPoint._Application app = new Microsoft.Office.Interop.PowerPoint.Application();
                app.Visible = MsoTriState.msoTrue;
                Microsoft.Office.Interop.PowerPoint.Presentation ppt = app.Presentations.Open(filepath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);


                //设置水印文字和字体
                System.Drawing.Font font = new System.Drawing.Font("宋体", 10);
                String watermark = "内部资料";

                foreach (Microsoft.Office.Interop.PowerPoint.Slide slide in ppt.Slides)
                {
                    //slide.DisplayMasterShapes = MsoTriState.msoCTrue;
                    slide.HeadersFooters.Footer.Visible = MsoTriState.msoCTrue;
                    slide.HeadersFooters.Footer.Text = watermark;
                }

                //ppt.SaveAs(@"C:\Users\Administrator\Desktop\Test_yemei.pptx",PpSaveAsFileType.ppSaveAsDefault);
                ppt.Save();
                ppt.Close();
                app.Quit();
                app = null;
                GC.Collect();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }
#endregion

        private void btPPT_Click(object sender, EventArgs e)
        {
            string filepath = "";
            string filetype = "";
            OpenFileDialog openFileDlg = new OpenFileDialog();
            if (openFileDlg.ShowDialog() == DialogResult.OK)
            {
                filepath = openFileDlg.FileName;
                filetype = Path.GetExtension(filepath);
                if (filetype != ".ppt" && filetype != ".pptx")
                {
                    MessageBox.Show("请选项PowerPoint类型的文件");
                    return;
                }
                tbPPT.Text = filepath;
                AddPageHeaderFooterForPPT(filepath);

            }
            
        }
    }
}

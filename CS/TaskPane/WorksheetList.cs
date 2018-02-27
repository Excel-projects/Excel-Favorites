using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Favorites.Scripts;
using Excel = Microsoft.Office.Interop.Excel;

namespace Favorites.TaskPane
{
    public partial class WorksheetList : UserControl
    {
        public WorksheetList()
        {
            InitializeComponent();
        }

        private void WorksheetList_Load(object sender, EventArgs e)
        {
            LoadListview();
        }

        private void lstWorksheets_SelectedIndexChanged(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets[lstWorksheets.SelectedItems[0].Text];
            sheet.Select(Type.Missing);
        }

        private void lstWorksheets_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Right)
                {
                    if (lstWorksheets.FocusedItem.Bounds.Contains(e.Location))
                    {
                        mnuSetVisiblity.Show(Cursor.Position);
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        private void tmiVisiblity_Click(object sender, EventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets[lstWorksheets.SelectedItems[0].Text];
                switch (sender.ToString())
                {
                    case "Visible":
                        sheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible;
                        break;
                    case "Hidden":
                        sheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden;
                        break;
                    case "Very Hidden":
                        sheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVeryHidden;
                        break;
                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
            finally
            {
                LoadListview();
            }
        }

        private void LoadListview()
        {
            try
            {
                lstWorksheets.Items.Clear();
                Microsoft.Office.Interop.Excel.Workbook book = Globals.ThisAddIn.Application.ActiveWorkbook;
                lstWorksheets.SmallImageList = imgXlSheetVisibility;
                //lstWorksheets.Columns.Add("Sheet Name");
                int visible = 0;
                foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in book.Worksheets)
                {
                    switch (sheet.Visible)
                    {
                        case Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible:
                            visible = 0;
                            break;
                        case Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden:
                            visible = 1;
                            break;
                        case Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVeryHidden:
                            visible = 2;
                            break;
                    }

                    ListViewItem lvi = new ListViewItem();
                    lvi.Text = sheet.Name;
                    lvi.ImageIndex = visible;
                    lstWorksheets.Items.Add(lvi);

                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        private void tsbRefresh_Click(object sender, EventArgs e)
        {
            LoadListview();
        }

        private void tsbSortAsc_Click(object sender, EventArgs e)
        {
            System.Collections.ArrayList al = new System.Collections.ArrayList();
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            int x = 0;
            try
            {
                foreach (Excel.Worksheet s in wb.Sheets)
                {
                    al.Add(s.Name);
                }
                al.Sort();

                foreach (string item in al)
                {
                    Excel.Worksheet s = (Excel.Worksheet)wb.Sheets[item];
                    //s.Move(wb.Sheets[wb.Sheets.Count]);
                    s.Move(wb.Sheets[x+=1]);
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
            finally
            {
                LoadListview();
            }

        }

        private void tsbSortDesc_Click(object sender, EventArgs e)
        {
            System.Collections.ArrayList al = new System.Collections.ArrayList();
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            int x = 0;
            try
            {
                foreach (Excel.Worksheet s in wb.Sheets)
                {
                    al.Add(s.Name);
                }
                al.Sort();

                foreach (string item in al)
                {
                    Excel.Worksheet s = (Excel.Worksheet)wb.Sheets[item];
                    int y = wb.Sheets.Count - x;
                    s.Move(wb.Sheets[y]);
                    x += 1;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
            finally
            {
                LoadListview();
            }

        }
    }
}

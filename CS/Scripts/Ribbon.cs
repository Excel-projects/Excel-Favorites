using System;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace Favorites.Scripts
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        public static Ribbon ribbonref;
        public TaskPane.Settings mySettings;
        public Microsoft.Office.Tools.CustomTaskPane myTaskPaneSettings;
        public TaskPane.WorksheetList myWorksheetList;
        public Microsoft.Office.Tools.CustomTaskPane myTaskPaneWorksheetList;

        #region | Ribbon Events |

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Favorites.Ribbon.xml");
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            try
            {
                this.ribbon = ribbonUI;
                ribbonref = this;
                AssemblyInfo.SetAddRemoveProgramsIcon("ExcelAddin.ico");
                AssemblyInfo.SetAssemblyFolderVersion();

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        public System.Drawing.Bitmap GetButtonImage(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "btnProblemStepRecorder":
                        return Properties.Resources.problem_steps_recorder;
                    case "btnSnippingTool":
                        return Properties.Resources.snipping_tool;
                    default:
                        return null;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return null;
            }
        }

        public string GetLabelText(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "tabFavorites":
                        if (Application.ProductVersion.Substring(0, 2) == "15") //for Excel 2013
                        {
                            return AssemblyInfo.Title.ToUpper();
                        }
                        else
                        {
                            return AssemblyInfo.Title;
                        }
                    case "txtCopyright":
                        return "© " + AssemblyInfo.Copyright;
                    case "txtDescription":
                        return AssemblyInfo.Title.Replace("&", "&&") + " " + AssemblyInfo.AssemblyVersion;
                    case "txtReleaseDate":
                        DateTime dteCreateDate = Properties.Settings.Default.App_ReleaseDate;
                        return dteCreateDate.ToString("dd-MMM-yyyy hh:mm tt");
                    default:
                        return string.Empty;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return string.Empty;
            }
        }

        public void OnAction(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "btnSettings":
                        OpenSettings();
                        break;
                    case "btnSheetVisibility":
                        OpenSheetVisibility();
                        break;
                    case "btnSnippingTool":
                        OpenSnippingTool();
                        break;
                    case "btnProblemStepRecorder":
                        OpenProblemStepRecorder();
                        break;
                    case "btnOpenReadMe":
                        OpenReadMe();
                        break;
                    case "btnOpenNewIssue":
                        OpenNewIssue();
                        break;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }

        }

        #endregion

        #region | Ribbon Buttons |

        public void CopyVisibleCells()
        {
            Excel.Range visibleRange = null;
            try
            {
                if (ErrorHandler.IsEnabled(true) == false)
                {
                    return;
                }
                ErrorHandler.CreateLogRecord();
                visibleRange = Globals.ThisAddIn.Application.Selection.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                visibleRange.Copy();
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
            finally
            {
                if (visibleRange != null)
                    Marshal.ReleaseComObject(visibleRange);
            }
        }

        public void OpenNewIssue()
        {
            try
            {
                System.Diagnostics.Process.Start(Properties.Settings.Default.App_PathReportIssue);

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        public void OpenProblemStepRecorder()
        {
            string filePath = @"C:\Windows\System32\psr.exe";
            try
            {
                System.Diagnostics.Process.Start(filePath);

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        public void OpenReadMe()
        {
            try
            {
                System.Diagnostics.Process.Start(Properties.Settings.Default.App_PathReadMe);

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        public void OpenSheetVisibility()
        {
            try
            {
                if (myTaskPaneWorksheetList != null)
                {
                    if (myTaskPaneWorksheetList.Visible == true)
                    {
                        myTaskPaneWorksheetList.Visible = false;
                    }
                    else
                    {
                        myTaskPaneWorksheetList.Visible = true;
                    }
                }
                else
                {
                    myWorksheetList = new TaskPane.WorksheetList();
                    myTaskPaneWorksheetList = Globals.ThisAddIn.CustomTaskPanes.Add(myWorksheetList, "Worksheet List");
                    myTaskPaneWorksheetList.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                    myTaskPaneWorksheetList.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
                    myTaskPaneWorksheetList.Width = 300;
                    myTaskPaneWorksheetList.Visible = true;
                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        public void OpenSettings()
        {
            try
            {
                if (myTaskPaneSettings != null)
                {
                    if (myTaskPaneSettings.Visible == true)
                    {
                        myTaskPaneSettings.Visible = false;
                    }
                    else
                    {
                        myTaskPaneSettings.Visible = true;
                    }
                }
                else
                {
                    mySettings = new TaskPane.Settings();
                    myTaskPaneSettings = Globals.ThisAddIn.CustomTaskPanes.Add(mySettings, "Settings for " + Scripts.AssemblyInfo.Title);
                    myTaskPaneSettings.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                    myTaskPaneSettings.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
                    myTaskPaneSettings.Width = 675;
                    myTaskPaneSettings.Visible = true;
                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        public void OpenSnippingTool()
        {
            string filePath;
            try
            {
                if (System.Environment.Is64BitOperatingSystem)
                {
                    filePath = "C:\\Windows\\sysnative\\SnippingTool.exe";
                }
                else
                {
                    filePath = "C:\\Windows\\system32\\SnippingTool.exe";
                }
                System.Diagnostics.Process.Start(filePath);

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        #endregion

    }

}

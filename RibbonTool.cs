using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using System.Windows.Forms;
using Autodesk.Windows;
using System.Windows.Media.Imaging;
using System.IO;
using System.Drawing;
using System.Windows.Input;

using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace AutoCadPlugIn
{
    public class RibbonTool
    {
        public static RibbonTab fabtab;


        public void CreateFabTab()
        {
            try
            {
                Autodesk.Windows.RibbonControl ribbonControl = ComponentManager.Ribbon;
                if (ribbonControl != null)
                {
                    fabtab = ribbonControl.FindTab("Properties Update");
                    if (fabtab == null)
                    {
                        fabtab = new RibbonTab();
                        fabtab.Title = "Properties Update";
                        fabtab.Id = "Properties Update";

                        ribbonControl.Tabs.Add(fabtab);
                        fabtab.IsActive = true;
                    }



                    //Fab panel
                    CreateFabPanel();
                }
            }
            catch (System.Exception ex)
            {
                throw;

            }
        }

        private void CreateFabPanel()
        {
            try
            {
                RibbonPanel existingPanel = null;
                foreach (RibbonPanel item in fabtab.Panels)
                {
                    if (item.UID == "Properties Update")
                    {
                        existingPanel = item;
                        break;
                    }
                }

                if (existingPanel != null)
                {
                    List<Autodesk.Windows.RibbonItem> RibbonItemLat = new List<Autodesk.Windows.RibbonItem>();
                    foreach (Autodesk.Windows.RibbonItem item in existingPanel.Source.Items)
                    {
                        RibbonItemLat.Add(item);
                    }
                    existingPanel.Source.Items.Clear();
                    existingPanel.Source.Items.Add(CreateSenderRibbonButton("Signal", "Click to update signal properties", "Signal"));
                    foreach (Autodesk.Windows.RibbonItem item in RibbonItemLat)
                    {
                        existingPanel.Source.Items.Add(item);
                    }
                }
                else
                {
                    RibbonPanel Panel1 = new RibbonPanel();
                    Panel1.UID = "Properties Update";
                    Autodesk.Windows.RibbonPanelSource projectsPanel = new Autodesk.Windows.RibbonPanelSource();
                    projectsPanel.Title = "Properties Update";
                    projectsPanel.Id = "Properties Update";
                    Panel1.Source = projectsPanel;

                    fabtab.Panels.Add(Panel1);
                    projectsPanel.Items.Add(CreateSenderRibbonButton("Signal", "Click to update signal properties", "Signal"));
                }
            }
            catch (System.Exception ex)
            {
                throw;
            }
        }



        /// <summary>
        /// Create ribbon button
        /// </summary>
        /// <param name="btnDetails">Ribbon button details</param>
        /// <returns>Ribbon Item</returns>
        private Autodesk.Windows.RibbonItem CreateSenderRibbonButton(string btnName, string btnToolTip, string btnCommand)
        {
            Autodesk.Windows.RibbonButton ribButton = new Autodesk.Windows.RibbonButton();
            try
            {
                ribButton.Text = btnName;
                ribButton.ToolTip = btnToolTip;
                ribButton.Size = RibbonItemSize.Large;
                ribButton.Orientation = System.Windows.Controls.Orientation.Vertical;
                ribButton.ShowText = true;
                ribButton.Id = btnToolTip;
                if (btnCommand != string.Empty)
                {
                    ribButton.CommandHandler = new Executer(btnCommand);
                }

            }
            catch (System.Exception ex)
            {
                throw;
            }
            return ribButton;
        }

        private BitmapImage Bitmap2BitmapImage(object senderIcon)
        {
            throw new NotImplementedException();
        }


    }



    internal class Executer : ICommand
    {
        private string btnCommand;

        Dictionary<string, string> Signal1 = new Dictionary<string, string>();
        Dictionary<string, string> Signal2 = new Dictionary<string, string>();

        public Executer(string btnCommand)
        {
            this.btnCommand = btnCommand;
        }

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
        }
        public void Execute(object parameter)
        {
            try
            {
                if (!string.IsNullOrEmpty(btnCommand))
                {
                    GetValuesfromExcel();
                    if (Signal1.Count > 0)
                        Properties_Update.PsetProperty.CreatePSetTab(Signal1, "Signal 1");
                    if (Signal2.Count > 0)
                        Properties_Update.PsetProperty.CreatePSetTab(Signal2, "Signal 2");
                    MessageBox.Show("Process Completed !!");
                }
            }
            catch (System.Exception ex)
            {

            }
        }

        /// <summary>
        ///Read Excel
        /// </summary>
        private void GetValuesfromExcel()
        {

            try
            {
                OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file
                if (file.ShowDialog() == DialogResult.OK) //if there is a file chosen by the user
                {
                    string fileExt = Path.GetExtension(file.FileName); //get the file extension
                    if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                    {
                        try
                        {
                            Excel.Application xlApp = new Excel.Application();
                            Excel.Workbook excelWorkbook = xlApp.Workbooks.Open(file.FileName, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                            Excel.Worksheet sheet = excelWorkbook.ActiveSheet;
                            Range xlRange = sheet.UsedRange;
                            for (int i = 2; i <= xlRange.Count + 1; i++)
                            {
                                var Key = (xlRange.Cells[1, i] as Excel.Range).Value2;
                                var value = (xlRange.Cells[2, i] as Excel.Range).Value2;
                                if (!string.IsNullOrEmpty(Convert.ToString(Key)) && !string.IsNullOrEmpty(Convert.ToString(value)))
                                    Signal1.Add(Convert.ToString(Key), Convert.ToString(value));
                            }
                            for (int i = 2; i <= xlRange.Count + 1; i++)
                            {
                                var Key = (xlRange.Cells[1, i] as Excel.Range).Value2;
                                var value = (xlRange.Cells[3, i] as Excel.Range).Value2;
                                if (!string.IsNullOrEmpty(Convert.ToString(Key)) && !string.IsNullOrEmpty(Convert.ToString(value)))
                                    Signal2.Add(Convert.ToString(Key), Convert.ToString(value));
                            }
                            excelWorkbook.Close();
                            xlApp.Quit();
                            Marshal.ReleaseComObject(sheet);
                            Marshal.ReleaseComObject(excelWorkbook);
                            Marshal.ReleaseComObject(xlApp);
                        }
                        catch (System.Exception ex)
                        {
                            //MessageBox.Show(ex.Message.ToString());
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error
                    }
                }
            }
            catch (System.Exception ex)
            {

            }
        }

        private bool UpdateSignalPropeties(string signalname)
        {
            bool isupdated = false;
            try
            {
                Document document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                using (Transaction trans = document.Database.TransactionManager.StartTransaction())
                {
                    DocumentLock LckDoc = document.LockDocument();

                    //PropertySet pset;
                    //BlockTable bt = trans.GetObject(document.Database.BlockTableId, OpenMode.ForWrite) as BlockTable;
                    //BlockTableRecord btr = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord; 
                    //foreach (ObjectId item in btr)
                    //{
                    //    DBObject obj = (DBObject)trans.GetObject(item, OpenMode.ForWrite);
                    //    BlockReference bref = obj as BlockReference;
                    //    if (bref != null)
                    //    {
                    //        DynamicBlockReferencePropertyCollection coll = bref.DynamicBlockReferencePropertyCollection;
                    //        foreach (ObjectId objid in coll)
                    //        {
                    //            AttributeReference attRef = (AttributeReference)trans.GetObject(objid, OpenMode.ForRead);

                    //        }
                    //    }
                    //}
                }
            }
            catch (System.Exception ex)
            {

            }
            return isupdated;
        }
    }
}
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;

namespace PPT_POC.Ribbon
    {
    [ComVisible(true)]
    public class PPTRibbon : Office.IRibbonExtensibility
        {
        private Office.IRibbonUI ribbon;

        public PPTRibbon() { }

        public void OnSerializeChartButton(Office.IRibbonControl control)
            {
            try
                {
                PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
                PowerPoint.Presentation presentation = pptApp.ActivePresentation;
                PowerPoint.Slide slide = pptApp.ActiveWindow.View.Slide;

                if (slide.Shapes.Count > 0)
                    {
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                        {
                        if (shape.HasChart == Office.MsoTriState.msoTrue)
                            {
                            PowerPoint.ChartData chartData = shape.Chart.ChartData;
                            chartData.Activate();
                            Excel.Workbook workbook = (Excel.Workbook)chartData.Workbook;

                            // Ensure Excel Add-in is loaded
                            LoadExcelAddIn(workbook.Application);

                            SelectLinkedChartInWorkbook(workbook, shape);
                            }
                        }
                    }
                else
                    {
                    MessageBox.Show("No chart selected.");
                    }
                }
            catch (Exception ex)
                {
                MessageBox.Show($"Error: {ex.Message}");
                }
            }

        private void SelectLinkedChartInWorkbook(Excel.Workbook workbook, PowerPoint.Shape pptChartShape)
            {
            try
                {
                PowerPoint.Chart pptChart = pptChartShape.Chart;
                var pptChartData = GetChartSeriesData(pptChart);

                bool chartFound = false;

                foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                    {
                    Excel.ChartObjects chartObjects = (Excel.ChartObjects)worksheet.ChartObjects();

                    for (int i = 1; i <= chartObjects.Count; i++)
                        {
                        Excel.ChartObject chartObject = chartObjects.Item(i);
                        var excelChartData = GetChartSeriesData(chartObject.Chart);

                        if (ChartDataMatches(pptChartData, excelChartData))
                            {
                            chartObject.Select();
                            worksheet.Activate();

                            chartFound = true;
                            // Get running Excel application
                            Excel.Application excelApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

                            // Get the Excel VSTO add-in


                            Office.COMAddIn excelAddIn = excelApp.COMAddIns.Item("Excel-POC");


                            // Unload the add-in
                            excelAddIn.Connect = false;
                            System.Threading.Thread.Sleep(1000); // Small delay to ensure it's fully unloaded

                            // Load the add-in again (this triggers ThisAddIn_Startup)
                            excelAddIn.Connect = true;
                            break;
                            }
                        }

                    if (chartFound) break;
                    }

                if (!chartFound)
                    {
                    MessageBox.Show("Linked chart not found in the Excel workbook across all sheets.");
                    }
                }
            catch (Exception ex)
                {
                MessageBox.Show($"Error selecting linked chart: {ex.Message}");
                }
            }

        private void LoadExcelAddIn(Excel.Application excelApp)
            {
            try
                {
                foreach (Excel.AddIn addIn in excelApp.AddIns)
                    {
                    if (!addIn.Installed)
                        {
                        addIn.Installed = true;
                        }
                    }
                }
            catch (Exception ex)
                {
                MessageBox.Show($"Error loading Excel Add-in: {ex.Message}");
                }
            }

        private Dictionary<string, List<double>> GetChartSeriesData(dynamic chart)
            {
            var seriesData = new Dictionary<string, List<double>>();

            foreach (dynamic series in chart.SeriesCollection())
                {
                string seriesName = series.Name;
                List<double> values = new List<double>();

                foreach (var value in series.Values as Array)
                    {
                    if (double.TryParse(value.ToString(), out double number))
                        {
                        values.Add(number);
                        }
                    }

                seriesData[seriesName] = values;
                }

            return seriesData;
            }

        private bool ChartDataMatches(Dictionary<string, List<double>> pptChartData, Dictionary<string, List<double>> excelChartData)
            {
            if (pptChartData.Count != excelChartData.Count)
                return false;

            foreach (var series in pptChartData)
                {
                if (!excelChartData.ContainsKey(series.Key))
                    return false;

                if (!series.Value.SequenceEqual(excelChartData[series.Key]))
                    return false;
                }

            return true;
            }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
            {
            return GetResourceText("PPT_POC.Ribbon.PPTRibbon.xml");
            }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
            {
            this.ribbon = ribbonUI;
            }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
            {
            Assembly asm = Assembly.GetExecutingAssembly();
            using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceName)))
                {
                return resourceReader?.ReadToEnd();
                }
            }

        #endregion
        }
    }

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_POC.Helpers
    {
    public static class ExcelChartHelper
        {
        public static Excel.Chart GetActiveChart()
            {
            try
                {
                Excel.Application excelApp = Globals.ThisAddIn.Application;
                Excel.Chart chart = excelApp.ActiveChart;
                return chart;
                }
            catch (Exception ex)
                {
                ErrorHandler.ShowError($"Error retrieving the active chart: {ex.Message}");
                return null;
                }
            }

        public static string BuildChartData(Excel.Chart chart)
            {
            try
                {
                Excel.SeriesCollection seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection();
                string csvData = "";
                foreach (Excel.Series series in seriesCollection)
                    {
                    csvData += GetSeriesData(series);
                    }
                return csvData;
                }
            catch (Exception ex)
                {
                ErrorHandler.ShowError($"Error building chart data: {ex.Message}");
                return string.Empty;
                }
            }

        private static string GetSeriesData(Excel.Series series)
            {
            try
                {
                string seriesName = series.Name;
                dynamic xValues = series.XValues;
                dynamic yValues = series.Values;
                object[] xArray = xValues as object[] ?? (xValues as System.Array)?.OfType<object>().ToArray();
                object[] yArray = yValues as object[] ?? (yValues as System.Array)?.OfType<object>().ToArray();

                int length = Math.Min(xArray.Length, yArray.Length);
                string seriesData = "";
                for (int i = 0; i < length; i++)
                    {
                    seriesData += $"{xArray[i]}, {yArray[i]}\n";
                    }
                return seriesData;
                }
            catch (Exception ex)
                {
                ErrorHandler.ShowError($"Error retrieving series data: {ex.Message}");
                return string.Empty;
                }
            }
        }
    }

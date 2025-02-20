using Excel_POC.Helpers;
using System;
using Excel_POC.Services;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_POC.TaskPane
    {
    /// <summary>
    /// Interaction logic for HomeTaskPane.xaml
    /// </summary>
    public partial class HomeTaskPane : UserControl
        {
        private string csvData;
        public HomeTaskPane()
            {
            InitializeComponent();
            GetSelectedChartData();
            }

        /// <summary>
        /// Extracts data from the selected Excel chart and displays it in the TaskPane TextBox.
        /// </summary>
        public void GetSelectedChartData()
            {
            try
                {
                Excel.Chart chart = ExcelChartHelper.GetActiveChart();
                if (chart == null) return;

                 csvData = ExcelChartHelper.BuildChartData(chart);

                // Update UI on the dispatcher thread
                SerializedDataTextBox.Dispatcher.Invoke(() =>
                {
                    SerializedDataTextBox.Text = csvData;
                });
                }
            catch (Exception ex)
                {
                ErrorHandler.ShowError($"An unexpected error occurred: {ex.Message}");
                }
            }




        /// <summary>
        /// Button click event handler to submit selected chart data to external api.
        /// </summary>
        private async void SubmitChartDataButton_Click(object sender, RoutedEventArgs e)
            {
            try
                {
              
                if (string.IsNullOrWhiteSpace(csvData))
                    {
                    MessageBox.Show("No data to send to the API.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                    }

                string response = await SaveDataApi.SendDataToApiAsync(csvData);
                MessageBox.Show($"API Response: {response}", "API Response", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            catch (Exception ex)
                {
                ErrorHandler.ShowError($"An error occurred while sending data to the API: {ex.Message}");
                }
            }
        }
    }
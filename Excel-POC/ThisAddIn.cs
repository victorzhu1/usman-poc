using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Excel_POC.TaskPane;
using Excel_POC.Ribbon;

namespace Excel_POC
    {
    public partial class ThisAddIn
        {

        // Event handler for Add-In startup
        private void ThisAddIn_Startup(object sender, EventArgs e)
            {
            ExcelRibbon excelRibbon = new ExcelRibbon();
            excelRibbon.LaunchAddIn();
            }

        // Method to create the ribbon extensibility object
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
            {
            try
                {
                return new ExcelRibbon();  // Return an instance of the Excel ribbon
                }
            catch (Exception ex)
                {
                // Display an error message if ribbon creation fails
                MessageBox.Show($"Error creating ribbon extensibility object: {ex.Message}", "Add-In Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
                }
            }

        // Event handler for Add-In shutdown
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
            {
            try
                {
                // Cleanup code if needed before the Add-In shuts down
                }
            catch (Exception ex)
                {
                // Display an error message if shutdown fails
                MessageBox.Show($"Error during Add-In shutdown: {ex.Message}", "Add-In Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        #region VSTO generated code

        // Internal startup method that hooks up the startup and shutdown events
        private void InternalStartup()
            {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
            }

        #endregion
        }
    }
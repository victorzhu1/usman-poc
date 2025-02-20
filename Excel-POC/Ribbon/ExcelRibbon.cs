using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms; // For MessageBox
using Office = Microsoft.Office.Core;
using Excel_POC.TaskPane;
using Microsoft.Office.Core;

namespace Excel_POC.Ribbon
    {
    [ComVisible(true)]
    public class ExcelRibbon : Office.IRibbonExtensibility
        {
        private Office.IRibbonUI ribbon;
        private Utilities.InitTaskPane taskPane;

        // Constructor
        public ExcelRibbon() { }

        /// <summary>
        /// Callback for the button click event to serialize chart data.
        /// </summary>
        /// <param name="control">The Ribbon control that triggered the event.</param>
        public void OnSerializeChartButton(Office.IRibbonControl control)
            {
            try
                {

                LaunchAddIn();
                }
            catch (Exception ex)
                {
                ShowErrorMessage($"An error occurred while serializing chart data: {ex.Message}");
                }
            }
        public void LaunchAddIn()
            {
            if (taskPane == null)
                {
                taskPane = new Utilities.InitTaskPane();
                taskPane.InitializeTaskPane();
                }


            }
        #region IRibbonExtensibility Members

        /// <summary>
        /// Loads the Ribbon XML.
        /// </summary>
        /// <param name="ribbonID">The ID of the ribbon.</param>
        /// <returns>Ribbon XML as a string.</returns>
        public string GetCustomUI(string ribbonID)
            {
            try
                {
                return GetResourceText("Excel_POC.Ribbon.ExcelRibbon.xml");
                }
            catch (Exception ex)
                {
                ShowErrorMessage($"Failed to load Ribbon XML: {ex.Message}");
                return null;
                }
            }

        #endregion

        #region Ribbon Callbacks

        /// <summary>
        /// Ribbon load event.
        /// </summary>
        /// <param name="ribbonUI">The Ribbon UI instance.</param>
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
            {
            this.ribbon = ribbonUI;
            }

        #endregion

        #region Helpers

        /// <summary>
        /// Retrieves the embedded resource text.
        /// </summary>
        /// <param name="resourceName">The name of the resource to retrieve.</param>
        /// <returns>The resource text, or null if not found.</returns>
        private static string GetResourceText(string resourceName)
            {
            try
                {
                Assembly asm = Assembly.GetExecutingAssembly();
                string[] resourceNames = asm.GetManifestResourceNames();

                foreach (string resource in resourceNames)
                    {
                    if (string.Equals(resourceName, resource, StringComparison.OrdinalIgnoreCase))
                        {
                        using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resource)))
                            {
                            return resourceReader?.ReadToEnd();
                            }
                        }
                    }
                ShowErrorMessage($"Resource not found: {resourceName}");
                }
            catch (Exception ex)
                {
                ShowErrorMessage($"An error occurred while retrieving resource text: {ex.Message}");
                }
            return null;
            }

        /// <summary>
        /// Shows an error message in a MessageBox.
        /// </summary>
        /// <param name="message">The error message to display.</param>
        private static void ShowErrorMessage(string message)
            {
            MessageBox.Show(message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        #endregion
        }
    }

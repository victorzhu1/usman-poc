using System;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Office = Microsoft.Office.Core;
using Excel_POC.TaskPane;

namespace Excel_POC.Utilities
    {
    /// <summary>
    /// Utility class to initialize and manage the custom task pane in Excel.
    /// </summary>
    public class InitTaskPane
        {
        private CustomTaskPane customTaskPane;
        private TaskPaneHost mediaTaskPane;

        /// <summary>
        /// Initializes and displays the custom task pane.
        /// </summary>
        public void InitializeTaskPane()
            {
            try
                {
                // Create an instance of the TaskPaneHost (WPF UserControl)
                mediaTaskPane = new TaskPaneHost();

                // Add the task pane to the Excel Add-in's custom task panes collection
                customTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(mediaTaskPane, "Excel POC");

                // Set the docking position of the task pane (Right side of the Excel window)
                customTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;

                // Set the width of the task pane
                customTaskPane.Width = 500;

                // Make the task pane visible
                customTaskPane.Visible = true;
                }
            catch (Exception ex)
                {
                // Show a message box with the error information if initialization fails
                MessageBox.Show($"An error occurred while initializing the task pane: {ex.Message}",
                                "Task Pane Initialization Error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                }
            }
        }
    }

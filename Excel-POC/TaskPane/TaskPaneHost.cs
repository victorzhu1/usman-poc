using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel_POC.TaskPane
    {
    public partial class TaskPaneHost : UserControl
        {
        public TaskPaneHost()
            {
            InitializeComponent();
            var elementHost = new System.Windows.Forms.Integration.ElementHost();
            elementHost.Dock = DockStyle.Fill;
            elementHost.Child = new HomeTaskPane(); // Replace with your WPF UserControl

            // Add the ElementHost to the Windows Forms UserControl
            Controls.Add(elementHost);
            }
        }
    }

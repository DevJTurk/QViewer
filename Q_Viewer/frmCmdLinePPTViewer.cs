using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Q_Viewer
{
    public partial class frmCmdLinePPTViewer : Form
    {
        public frmCmdLinePPTViewer()
        {
            InitializeComponent();
        }

        private void frmCmdLinePPTViewer_Load(object sender, EventArgs e)
        {
            string installerPath = Path.Combine(Application.StartupPath, "Tools", "PowerPointViewer.exe");

            try
            {
                // Start the installer silently
                Process.Start(new ProcessStartInfo
                {
                    FileName = installerPath,
                    Arguments = "/quiet",  // Adjust command line switches as needed
                    CreateNoWindow = true,
                    UseShellExecute = false
                }).WaitForExit();

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error installing PowerPoint Viewer: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


            string powerPointViewerPath = @"C:\Program Files (x86)\Microsoft Office\Office14\PPTVIEW.exe";
            string presentationFilePath = Path.Combine(Application.StartupPath, "Presentation", "e_Capital_01_BACKUP1.ppsx");

            string commandLine = $@"/S /F""{presentationFilePath}""";

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = powerPointViewerPath,
                    Arguments = $"/C {commandLine}",
                    CreateNoWindow = true,
                    UseShellExecute = false
                });
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error starting PowerPoint Viewer: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Close();
            }
        }
    }
}

using ScalesData.Properties;
using System;
using System.Windows.Forms;

namespace ScalesData
{
    public partial class ExportExcelForm : Form
    {
        Launcher mLauncher;

        public ExportExcelForm(Launcher launcher)
        {
            mLauncher = launcher;
            InitializeComponent();

            this.textBoxDestination.DataBindings.Add(
                new Binding("Text", Settings.Default, "Destination", true));
            this.textBoxDbPath.DataBindings.Add(
                new Binding("Text", Settings.Default, "DataPath", true));
            this.checkBoxProtect.CheckedChanged += CheckBoxProtect_CheckedChanged;
        }

        private void CheckBoxProtect_CheckedChanged(object sender, EventArgs e)
        {
            this.textBoxPassword.Enabled = checkBoxProtect.Checked;
        }

        private void onExportClick(object sender, EventArgs e)
        {
            string startDate = null;
            string endDate = null;

            if (this.startDatePicker.Checked)
            {
                startDate = startDatePicker.Value.ToString();
            }
            if (this.endDatePicker.Checked)
            {
                endDate = endDatePicker.Value.ToString();
            }

            mLauncher.exportExcel(startDate, endDate, (checkBoxProtect.Checked) ? textBoxPassword.Text : null);
            //MessageBoxEx.Show(this, "Report Created.\t\t");
        }

        private void openSelectDestination(object sender, EventArgs e)
        {
            if (sender == buttonDataBrowse)
            {
                showSelectDatabaseDialog(false);
                return;
            }

            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (!string.IsNullOrWhiteSpace(folderBrowserDialog1.SelectedPath))
            {
                if (sender == buttonBrowse)
                {
                    Settings.Default.Destination = folderBrowserDialog1.SelectedPath;
                    Settings.Default.Save();
                }
            }
        }

        public static void showSelectDatabaseDialog(bool closeIfInvalid)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            DialogResult result = dialog.ShowDialog();
            string errorMsg = null;

            if (DataStorage.validateDataPath(dialog.SelectedPath, out errorMsg))
            {
                Settings.Default.DataPath = dialog.SelectedPath;
                Settings.Default.Save();
            } else
            {
                if (closeIfInvalid) errorMsg += " Application will be exited.";
                MessageBoxEx.Show(errorMsg);

                //if (result == DialogResult.Cancel)
                //{
                    if (closeIfInvalid)
                    {
                        Launcher launcher = new Launcher();
                        launcher.exit();
                    } else
                    {
                        if (!DataStorage.validateDataPath(Settings.Default.DataPath, out errorMsg))
                            showSelectDatabaseDialog(closeIfInvalid);
                    }
                //}
            }
        }
    }
}

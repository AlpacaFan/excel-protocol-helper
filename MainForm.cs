using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;


using ExcelWorkbook = Microsoft.Office.Interop.Excel.Workbook;

namespace ExcelProtocolHelper
{
    public partial class MainForm : Form
    {
        const string ProtocolName = "exceldata";

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {

            string[] arguments = Environment.GetCommandLineArgs();
            if(arguments.Length > 1)
            {
                string url = arguments[1];
                if(!url.StartsWith(ProtocolName))
                {
                    MessageBox.Show(this, "Invalid arguments to program! Must be an excel protocol URL", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    urlTextBox.Text = "http" + url.Substring(ProtocolName.Length);
                }
            }

            IList<ExcelWorkbook> workbooks = ExcelInterfaceUtility.GetAllOpenWorkbooks();

            workbookListBox.Items.Add(new ExcelWorkbookListItem(null, "New Excel Workbook"));
            foreach (ExcelWorkbook workbook in workbooks)
            {
                workbookListBox.Items.Add(new ExcelWorkbookListItem(workbook,null));
            }    
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
           Close();
        }

        private void addButton_Click(object sender, EventArgs e)
        {
            

            ExcelWorkbookListItem listItem = workbookListBox.SelectedItem as ExcelWorkbookListItem;
            if (listItem is null)
                return;

            ExcelInterfaceUtility.OpenLinkAsSheet(listItem.Workbook, this.urlTextBox.Text);
            workbookListBox.Items.Clear();
            this.Close();
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced);
            GC.WaitForPendingFinalizers();
        }

        private void CreateProtocolKeys(string protocol)
        {
            string protocolKey = "HKEY_CURRENT_USER\\Software\\Classes\\" + protocol;
            string commandKey = protocolKey + "\\shell\\open\\command";
            Registry.SetValue(protocolKey, null,"URL:Excel Protocol Helper", RegistryValueKind.String);
            Registry.SetValue(protocolKey, "URL Protocol","", RegistryValueKind.String);
            Registry.SetValue(commandKey, null, Environment.GetCommandLineArgs()[0]+" %1",RegistryValueKind.String);
        }

        private void registerButton_Click(object sender, EventArgs e)
        {
            CreateProtocolKeys(ProtocolName);
            CreateProtocolKeys(ProtocolName+"s");
        }
    }
}

using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExportDataGridViewToExcel
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            List<Product> liste = ((DataParameter)e.Argument).ProductList;
            string nomFichier = ((DataParameter)e.Argument).nomFichier;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = excel.Workbooks.Add(XlSheetType.xlWorksheet);
            Worksheet ws = (excel).ActiveSheet;
            excel.Visible = false;
            int index = 1;
            int process = liste.Count;

            //Add column
            ws.Cells[1, 1] = "ProductID";
            ws.Cells[1, 2] = "Product Name";
            ws.Cells[1, 3] = "Price";
            ws.Cells[1, 4] = "Stock";
            foreach (Product p in liste)
            {
                if(!backgroundWorker.CancellationPending)
                {
                    backgroundWorker.ReportProgress(index++ * 100 / process);
                    ws.Cells[index, 1] = p.CategoryID.ToString();
                    ws.Cells[index, 2] = p.ProductName;
                    ws.Cells[index, 3] = p.UnitPrice.ToString();
                    ws.Cells[index, 4] = p.UnitsInStock.ToString();
                }
            }
            ws.SaveAs(nomFichier, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
            excel.Quit();
        }
        struct DataParameter
        {
            public List<Product> ProductList;
            public string nomFichier { get; set; }
        }
        DataParameter _inputParameter;
        private void frmMain_Load(object sender, EventArgs e)
        {
            using(NorthwindEntities db  = new NorthwindEntities())
            {
                productBindingSource.DataSource = db.Products.ToList();
            }
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
            lblEtat.Text = string.Format("Traitement...{0}", e.ProgressPercentage);
            progressBar.Update();
        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error==null)
            {
                Thread.Sleep(100);
                lblEtat.Text = "Vos données ont été exportées avec succès !";
            }
        }

        private void btnExporter_Click(object sender, EventArgs e)
        {
            if (backgroundWorker.IsBusy)
                return;
            using(SaveFileDialog sfd = new SaveFileDialog() { Filter = "Fichiel Excel |*.xls" })
            {
                if(sfd.ShowDialog() == DialogResult.OK)
                {
                    _inputParameter.nomFichier = sfd.FileName;
                    _inputParameter.ProductList = productBindingSource.DataSource as List<Product>;
                    progressBar.Minimum = 0;
                    progressBar.Value = 0;
                    backgroundWorker.RunWorkerAsync(_inputParameter);
                }
            }
        }
    }
}

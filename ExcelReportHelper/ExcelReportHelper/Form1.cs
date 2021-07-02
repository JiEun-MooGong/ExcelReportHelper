using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelReportHelper
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            try
            {
                DataRow dr = ((DataRowView)grid.CurrentRow.DataBoundItem).Row;
                ExcelReportHelper ERHelper = new ExcelReportHelper(txtKeyword.Text, dr);
                ERHelper.Print();
            }
            catch ( Exception ex)
            {
                MessageBox.Show(ex.Message,"오류");
            }
        }
    }
}

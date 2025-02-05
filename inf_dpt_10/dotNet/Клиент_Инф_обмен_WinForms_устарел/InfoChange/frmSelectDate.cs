using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace InfoChange
{
    public partial class frmSelectDate : Form
    {
        public DatePeriod DateStartEnd;

        public frmSelectDate()
        {
            InitializeComponent();
            DateStartEnd.DateStart = DateTime.Today;
            DateStartEnd.DateEnd = DateTime.Today;
        }

        public DatePeriod ShowForm()
        {
            this.ShowDialog();
            return DateStartEnd;
        }

        private void dtpDateStart_ValueChanged(object sender, EventArgs e)
        {
            DateStartEnd.DateStart = dtpDateStart.Value;
            if (dtpDateEnd.Value < dtpDateStart.Value) dtpDateEnd.Value = dtpDateStart.Value;
        }

        private void dtpDateEnd_ValueChanged(object sender, EventArgs e)
        {
            DateStartEnd.DateEnd = dtpDateEnd.Value;
            if (dtpDateStart.Value > dtpDateEnd.Value) dtpDateStart.Value = dtpDateEnd.Value;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DateStartEnd.DateStart = DateTime.MinValue;
            DateStartEnd.DateEnd = DateTime.MinValue;
            this.Close();
        }

        private void btnReport_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}
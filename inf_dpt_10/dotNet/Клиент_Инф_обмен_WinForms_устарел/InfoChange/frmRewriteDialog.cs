using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace InfoChange
{
    public partial class frmRewriteDialog : Form
    {
        private int iResult = 1;
        public frmRewriteDialog()
        {
            InitializeComponent();
        }

        public int ShowForm(){

            this.ShowDialog();
            return iResult;
        }

        private void lblText_Click(object sender, EventArgs e)
        {

        }

        private void btnAppend_Click(object sender, EventArgs e)
        {
         //2 - дописать
         //20 - дописать все
            if (cbxAll.Checked) iResult = 20;
            else iResult = 2;
            this.Close();
        }

        private void btnRewrite_Click(object sender, EventArgs e)
        {
            //3 - перезаписать
            //21 - перезаписать все
            if (cbxAll.Checked) iResult = 21;
            else iResult = 3;
            this.Close();
        }

        private void btnNo_Click(object sender, EventArgs e)
        {
            //4 - пропустить
            //22 - пропустить все, которые найдены
            if (cbxAll.Checked) iResult = 22;
            else iResult = 4;
            this.Close();
        }

        private void cbxAll_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void frmRewriteDialog_Load(object sender, EventArgs e)
        {

        }
    }
}
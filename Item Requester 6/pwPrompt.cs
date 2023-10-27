using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Item_Requester_6
{

    public partial class pwPrompt : Form
    {
        public string password = string.Empty;

        public pwPrompt()
        {
            InitializeComponent();
            tbxPassword.Focus();
            tbxPassword.AcceptsReturn = true;
        }

        private void btnOkay_Click(object sender, EventArgs e)
        {
            password = tbxPassword.Text.ToString();
            this.Close();
        }


    }
}


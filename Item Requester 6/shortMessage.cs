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
    public partial class shortMessage : Form
    {
        public string bodyText = string.Empty;
        public bool unlisted;
        public string[] storeAdd = new string[4];
        public int arrayIndex = 0;
        //public shortMessage(bool unlistedName, int storeIndex)
        //{
        //    InitializeComponent();
        //    unlisted = unlistedName;

        //    switch(storeIndex)
        //    {

        //        case 2:
        //            cbxPPAdd.Checked = true;
        //            cbxWildAdd.Checked = true;
        //            break;

        //        case 1:
        //            cbxWoodAdd.Checked = true;
        //            cbxTemeAdd.Checked = true;
        //            cbxWildAdd.Checked = true;
        //            break;

        //        default:
        //            break;
        //    }

        private void btnOkay_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }



}
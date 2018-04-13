﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TourQueryManager
{
    public partial class FrmAdminPage : Form
    {
        public FrmAdminPage()
        {
            InitializeComponent();
        }

        private void FrmAdminPage_Load(object sender, EventArgs e)
        {
            /* this will be used in case of loading page of Admin screen */
        }

        private void btnCreateUser_Click(object sender, EventArgs e)
        {
            /* load the new screen for creating new user */
            FrmEditUserPage frmEditUserPage = new FrmEditUserPage("User");
            Hide();
            frmEditUserPage.ShowDialog();
            Show();
        }

        private void btnCreateQuery_Click(object sender, EventArgs e)
        {
            /* load the new screen to create the new Query */
            FrmEditQueryPage frmQuery = new FrmEditQueryPage();
            Hide();
            frmQuery.ShowDialog();
            Show();
        }

        private void btnCreateClient_Click(object sender, EventArgs e)
        {
            /* load the new screen for creating new user */
            FrmEditUserPage frmEditUserPage = new FrmEditUserPage("Client");
            Hide();
            frmEditUserPage.ShowDialog();
            Show();
        }
    }
}

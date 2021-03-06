﻿namespace TourQueryManager
{
    partial class FrmLoginPage
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmLoginPage));
            this.cmbboxUsername = new System.Windows.Forms.ComboBox();
            this.lblUserName = new System.Windows.Forms.Label();
            this.lblPassword = new System.Windows.Forms.Label();
            this.txtboxPassword = new System.Windows.Forms.TextBox();
            this.btnLogin = new System.Windows.Forms.Button();
            this.PctrBoxLogo = new System.Windows.Forms.PictureBox();
            this.PctrBoxSetting = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.PctrBoxLogo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.PctrBoxSetting)).BeginInit();
            this.SuspendLayout();
            // 
            // cmbboxUsername
            // 
            this.cmbboxUsername.FormattingEnabled = true;
            this.cmbboxUsername.Location = new System.Drawing.Point(308, 83);
            this.cmbboxUsername.Name = "cmbboxUsername";
            this.cmbboxUsername.Size = new System.Drawing.Size(121, 21);
            this.cmbboxUsername.TabIndex = 0;
            this.cmbboxUsername.SelectedIndexChanged += new System.EventHandler(this.cmbboxUsername_SelectedIndexChanged);
            // 
            // lblUserName
            // 
            this.lblUserName.AutoSize = true;
            this.lblUserName.Location = new System.Drawing.Point(247, 83);
            this.lblUserName.Name = "lblUserName";
            this.lblUserName.Size = new System.Drawing.Size(55, 13);
            this.lblUserName.TabIndex = 1;
            this.lblUserName.Text = "Username";
            // 
            // lblPassword
            // 
            this.lblPassword.AutoSize = true;
            this.lblPassword.Location = new System.Drawing.Point(247, 118);
            this.lblPassword.Name = "lblPassword";
            this.lblPassword.Size = new System.Drawing.Size(53, 13);
            this.lblPassword.TabIndex = 2;
            this.lblPassword.Text = "Password";
            // 
            // txtboxPassword
            // 
            this.txtboxPassword.Location = new System.Drawing.Point(308, 118);
            this.txtboxPassword.Name = "txtboxPassword";
            this.txtboxPassword.PasswordChar = '*';
            this.txtboxPassword.Size = new System.Drawing.Size(121, 20);
            this.txtboxPassword.TabIndex = 3;
            this.txtboxPassword.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtboxPassword_KeyPress);
            // 
            // btnLogin
            // 
            this.btnLogin.Location = new System.Drawing.Point(308, 219);
            this.btnLogin.Name = "btnLogin";
            this.btnLogin.Size = new System.Drawing.Size(121, 23);
            this.btnLogin.TabIndex = 4;
            this.btnLogin.Text = "Login";
            this.btnLogin.UseVisualStyleBackColor = true;
            this.btnLogin.Click += new System.EventHandler(this.btnLogin_Click);
            // 
            // PctrBoxLogo
            // 
            this.PctrBoxLogo.Image = global::TourQueryManager.Properties.Resources.ExcursionHolidaysLogo;
            this.PctrBoxLogo.Location = new System.Drawing.Point(12, 12);
            this.PctrBoxLogo.Name = "PctrBoxLogo";
            this.PctrBoxLogo.Size = new System.Drawing.Size(229, 230);
            this.PctrBoxLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.PctrBoxLogo.TabIndex = 30;
            this.PctrBoxLogo.TabStop = false;
            // 
            // PctrBoxSetting
            // 
            this.PctrBoxSetting.Image = global::TourQueryManager.Properties.Resources.Setting;
            this.PctrBoxSetting.Location = new System.Drawing.Point(382, 12);
            this.PctrBoxSetting.Name = "PctrBoxSetting";
            this.PctrBoxSetting.Size = new System.Drawing.Size(47, 42);
            this.PctrBoxSetting.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.PctrBoxSetting.TabIndex = 31;
            this.PctrBoxSetting.TabStop = false;
            this.PctrBoxSetting.DoubleClick += new System.EventHandler(this.PctrBoxSetting_DoubleClick);
            // 
            // FrmLoginPage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(444, 258);
            this.Controls.Add(this.PctrBoxSetting);
            this.Controls.Add(this.PctrBoxLogo);
            this.Controls.Add(this.btnLogin);
            this.Controls.Add(this.txtboxPassword);
            this.Controls.Add(this.lblPassword);
            this.Controls.Add(this.lblUserName);
            this.Controls.Add(this.cmbboxUsername);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "FrmLoginPage";
            this.Text = "Login";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmLoginPage_FormClosing);
            this.Load += new System.EventHandler(this.FrmLoginPage_Load);
            ((System.ComponentModel.ISupportInitialize)(this.PctrBoxLogo)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.PctrBoxSetting)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cmbboxUsername;
        private System.Windows.Forms.Label lblUserName;
        private System.Windows.Forms.Label lblPassword;
        private System.Windows.Forms.TextBox txtboxPassword;
        private System.Windows.Forms.Button btnLogin;
        private System.Windows.Forms.PictureBox PctrBoxLogo;
        private System.Windows.Forms.PictureBox PctrBoxSetting;
    }
}
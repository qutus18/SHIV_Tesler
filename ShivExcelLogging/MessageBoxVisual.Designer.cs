namespace ShivExcelLogging
{
    partial class MessageBoxVisual
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MessageBoxVisual));
            this.btnCloseForm2 = new System.Windows.Forms.Button();
            this.bunifuElipse1 = new Bunifu.Framework.UI.BunifuElipse(this.components);
            this.SuspendLayout();
            // 
            // btnCloseForm2
            // 
            this.btnCloseForm2.BackColor = System.Drawing.Color.Transparent;
            this.btnCloseForm2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnCloseForm2.Cursor = System.Windows.Forms.Cursors.Default;
            this.btnCloseForm2.FlatAppearance.BorderSize = 0;
            this.btnCloseForm2.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnCloseForm2.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnCloseForm2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCloseForm2.Font = new System.Drawing.Font("Century Gothic", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCloseForm2.Image = ((System.Drawing.Image)(resources.GetObject("btnCloseForm2.Image")));
            this.btnCloseForm2.Location = new System.Drawing.Point(40, 31);
            this.btnCloseForm2.Name = "btnCloseForm2";
            this.btnCloseForm2.Size = new System.Drawing.Size(115, 114);
            this.btnCloseForm2.TabIndex = 0;
            this.btnCloseForm2.TabStop = false;
            this.btnCloseForm2.UseMnemonic = false;
            this.btnCloseForm2.UseVisualStyleBackColor = false;
            this.btnCloseForm2.Click += new System.EventHandler(this.btnCloseForm2_Click);
            // 
            // bunifuElipse1
            // 
            this.bunifuElipse1.ElipseRadius = 5;
            this.bunifuElipse1.TargetControl = null;
            // 
            // MessageBoxVisual
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.MediumSeaGreen;
            this.ClientSize = new System.Drawing.Size(199, 180);
            this.Controls.Add(this.btnCloseForm2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MessageBoxVisual";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "MessageBoxVisual";
            this.TopMost = true;
            this.TransparencyKey = System.Drawing.Color.White;
            this.Load += new System.EventHandler(this.MessageBoxVisual_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnCloseForm2;
        private Bunifu.Framework.UI.BunifuElipse bunifuElipse1;
    }
}
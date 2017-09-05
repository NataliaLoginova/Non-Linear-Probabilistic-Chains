namespace Non_Linear_Porabalistic_Chain_WinForm
{
    partial class Form1
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
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.methodToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.logisticGrowToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.linearLogariphGrowToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.loadToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.methodToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1050, 33);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.loadToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(65, 29);
            this.fileToolStripMenuItem.Text = "Файл";
            // 
            // methodToolStripMenuItem
            // 
            this.methodToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.logisticGrowToolStripMenuItem,
            this.linearLogariphGrowToolStripMenuItem});
            this.methodToolStripMenuItem.Name = "methodToolStripMenuItem";
            this.methodToolStripMenuItem.Size = new System.Drawing.Size(77, 29);
            this.methodToolStripMenuItem.Text = "Метод";
            // 
            // logisticGrowToolStripMenuItem
            // 
            this.logisticGrowToolStripMenuItem.Name = "logisticGrowToolStripMenuItem";
            this.logisticGrowToolStripMenuItem.Size = new System.Drawing.Size(365, 30);
            this.logisticGrowToolStripMenuItem.Text = "Логистический рост";
            this.logisticGrowToolStripMenuItem.Click += new System.EventHandler(this.logisticGrowToolStripMenuItem_Click);
            // 
            // linearLogariphGrowToolStripMenuItem
            // 
            this.linearLogariphGrowToolStripMenuItem.Name = "linearLogariphGrowToolStripMenuItem";
            this.linearLogariphGrowToolStripMenuItem.Size = new System.Drawing.Size(365, 30);
            this.linearLogariphGrowToolStripMenuItem.Text = "Линейно-логарифмический рост";
            // 
            // loadToolStripMenuItem
            // 
            this.loadToolStripMenuItem.Name = "loadToolStripMenuItem";
            this.loadToolStripMenuItem.Size = new System.Drawing.Size(210, 30);
            this.loadToolStripMenuItem.Text = "Загрузить...";
            this.loadToolStripMenuItem.Click += new System.EventHandler(this.loadToolStripMenuItem_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1050, 684);
            this.Controls.Add(this.menuStrip1);
            this.IsMdiContainer = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "Form1";
            this.Text = "Non-linear Probabalistic Chain";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.Paint += new System.Windows.Forms.PaintEventHandler(this.Form1_Paint);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem methodToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem logisticGrowToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem linearLogariphGrowToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem loadToolStripMenuItem;
    }
}


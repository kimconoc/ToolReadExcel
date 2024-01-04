namespace ToolReadExcel
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.btnExecuteForwardedData = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.pathExcel = new System.Windows.Forms.TextBox();
            this.strConnect = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btnExecuteDeleteData = new System.Windows.Forms.Button();
            this.txtProces = new System.Windows.Forms.Label();
            this.NubProces = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnExecuteForwardedData
            // 
            resources.ApplyResources(this.btnExecuteForwardedData, "btnExecuteForwardedData");
            this.btnExecuteForwardedData.Name = "btnExecuteForwardedData";
            this.btnExecuteForwardedData.UseVisualStyleBackColor = true;
            this.btnExecuteForwardedData.Click += new System.EventHandler(this.btnExecuteForwardedData_Click);
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.label1.Name = "label1";
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.label2.Name = "label2";
            // 
            // pathExcel
            // 
            resources.ApplyResources(this.pathExcel, "pathExcel");
            this.pathExcel.Name = "pathExcel";
            // 
            // strConnect
            // 
            resources.ApplyResources(this.strConnect, "strConnect");
            this.strConnect.Name = "strConnect";
            // 
            // label3
            // 
            resources.ApplyResources(this.label3, "label3");
            this.label3.Name = "label3";
            // 
            // btnExecuteDeleteData
            // 
            resources.ApplyResources(this.btnExecuteDeleteData, "btnExecuteDeleteData");
            this.btnExecuteDeleteData.Name = "btnExecuteDeleteData";
            this.btnExecuteDeleteData.UseVisualStyleBackColor = true;
            this.btnExecuteDeleteData.Click += new System.EventHandler(this.btnExecuteDeleteData_Click);
            // 
            // txtProces
            // 
            resources.ApplyResources(this.txtProces, "txtProces");
            this.txtProces.Name = "txtProces";
            // 
            // NubProces
            // 
            resources.ApplyResources(this.NubProces, "NubProces");
            this.NubProces.ForeColor = System.Drawing.Color.Blue;
            this.NubProces.Name = "NubProces";
            // 
            // Form1
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.NubProces);
            this.Controls.Add(this.txtProces);
            this.Controls.Add(this.btnExecuteDeleteData);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.strConnect);
            this.Controls.Add(this.pathExcel);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnExecuteForwardedData);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Button btnExecuteForwardedData;
        private Label label1;
        private Label label2;
        private TextBox pathExcel;
        private TextBox strConnect;
        private Label label3;
        private Button btnExecuteDeleteData;
        private Label txtProces;
        private Label NubProces;
    }
}
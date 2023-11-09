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
            this.btnSubmit = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.pathExcel = new System.Windows.Forms.TextBox();
            this.strConnect = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.tetNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CATEGCD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SETCMDCD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CUSTCD1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.KORMKS = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnExport = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // btnSubmit
            // 
            resources.ApplyResources(this.btnSubmit, "btnSubmit");
            this.btnSubmit.Name = "btnSubmit";
            this.btnSubmit.UseVisualStyleBackColor = true;
            this.btnSubmit.Click += new System.EventHandler(this.btnSubmit_Click);
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
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
            // dataGridView
            // 
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.tetNo,
            this.CATEGCD,
            this.SETCMDCD,
            this.CUSTCD1,
            this.KORMKS});
            resources.ApplyResources(this.dataGridView, "dataGridView");
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.ReadOnly = true;
            this.dataGridView.RowHeadersVisible = false;
            this.dataGridView.RowTemplate.Height = 25;
            // 
            // tetNo
            // 
            resources.ApplyResources(this.tetNo, "tetNo");
            this.tetNo.Name = "tetNo";
            this.tetNo.ReadOnly = true;
            // 
            // CATEGCD
            // 
            resources.ApplyResources(this.CATEGCD, "CATEGCD");
            this.CATEGCD.Name = "CATEGCD";
            this.CATEGCD.ReadOnly = true;
            // 
            // SETCMDCD
            // 
            resources.ApplyResources(this.SETCMDCD, "SETCMDCD");
            this.SETCMDCD.Name = "SETCMDCD";
            this.SETCMDCD.ReadOnly = true;
            // 
            // CUSTCD1
            // 
            resources.ApplyResources(this.CUSTCD1, "CUSTCD1");
            this.CUSTCD1.Name = "CUSTCD1";
            this.CUSTCD1.ReadOnly = true;
            // 
            // KORMKS
            // 
            resources.ApplyResources(this.KORMKS, "KORMKS");
            this.KORMKS.Name = "KORMKS";
            this.KORMKS.ReadOnly = true;
            // 
            // btnExport
            // 
            this.btnExport.AutoEllipsis = true;
            resources.ApplyResources(this.btnExport, "btnExport");
            this.btnExport.Name = "btnExport";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // Form1
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.dataGridView);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.strConnect);
            this.Controls.Add(this.pathExcel);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnSubmit);
            this.MaximizeBox = false;
            this.Name = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Button btnSubmit;
        private Label label1;
        private Label label2;
        private TextBox pathExcel;
        private TextBox strConnect;
        private Label label3;
        private DataGridView dataGridView;
        private DataGridViewTextBoxColumn tetNo;
        private DataGridViewTextBoxColumn CATEGCD;
        private DataGridViewTextBoxColumn SETCMDCD;
        private DataGridViewTextBoxColumn CUSTCD1;
        private DataGridViewTextBoxColumn KORMKS;
        private Button btnExport;
    }
}
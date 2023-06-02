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
            this.txtCmdcd = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtMdkikaku = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtUpdt = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // btnSubmit
            // 
            resources.ApplyResources(this.btnSubmit, "btnSubmit");
            this.btnSubmit.Name = "btnSubmit";
            this.btnSubmit.UseVisualStyleBackColor = true;
            this.btnSubmit.Click += new System.EventHandler(this.button1_Click);
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
            this.txtCmdcd,
            this.txtMdkikaku,
            this.txtUpdt});
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
            // txtCmdcd
            // 
            resources.ApplyResources(this.txtCmdcd, "txtCmdcd");
            this.txtCmdcd.Name = "txtCmdcd";
            this.txtCmdcd.ReadOnly = true;
            // 
            // txtMdkikaku
            // 
            resources.ApplyResources(this.txtMdkikaku, "txtMdkikaku");
            this.txtMdkikaku.Name = "txtMdkikaku";
            this.txtMdkikaku.ReadOnly = true;
            // 
            // txtUpdt
            // 
            resources.ApplyResources(this.txtUpdt, "txtUpdt");
            this.txtUpdt.Name = "txtUpdt";
            this.txtUpdt.ReadOnly = true;
            // 
            // Form1
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
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
        private DataGridViewTextBoxColumn txtCmdcd;
        private DataGridViewTextBoxColumn txtMdkikaku;
        private DataGridViewTextBoxColumn txtUpdt;
    }
}
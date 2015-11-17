namespace NewBTASProto
{
    partial class WorkOrderReps
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WorkOrderReps));
            this.reportViewer1 = new Microsoft.Reporting.WinForms.ReportViewer();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.toolStripCBWorkOrderStatus = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.toolStripCBCustomers = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.toolStripCBSerialNums = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // reportViewer1
            // 
            this.reportViewer1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.reportViewer1.Location = new System.Drawing.Point(12, 32);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.Size = new System.Drawing.Size(903, 666);
            this.reportViewer1.TabIndex = 3;
            this.reportViewer1.ZoomMode = Microsoft.Reporting.WinForms.ZoomMode.PageWidth;
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePicker1.Location = new System.Drawing.Point(635, 5);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(97, 20);
            this.dateTimePicker1.TabIndex = 4;
            this.dateTimePicker1.Value = new System.DateTime(2000, 1, 1, 0, 0, 0, 0);
            this.dateTimePicker1.ValueChanged += new System.EventHandler(this.dateTimePicker1_ValueChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(596, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(33, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "From:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(744, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(23, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "To:";
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.Checked = false;
            this.dateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePicker2.Location = new System.Drawing.Point(773, 5);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(97, 20);
            this.dateTimePicker2.TabIndex = 7;
            this.dateTimePicker2.ValueChanged += new System.EventHandler(this.dateTimePicker2_ValueChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 8);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(46, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "Status:  ";
            // 
            // toolStripCBWorkOrderStatus
            // 
            this.toolStripCBWorkOrderStatus.FormattingEnabled = true;
            this.toolStripCBWorkOrderStatus.Items.AddRange(new object[] {
            "",
            "Open",
            "Closed",
            "Active"});
            this.toolStripCBWorkOrderStatus.Location = new System.Drawing.Point(61, 5);
            this.toolStripCBWorkOrderStatus.Name = "toolStripCBWorkOrderStatus";
            this.toolStripCBWorkOrderStatus.Size = new System.Drawing.Size(121, 21);
            this.toolStripCBWorkOrderStatus.TabIndex = 9;
            this.toolStripCBWorkOrderStatus.SelectedIndexChanged += new System.EventHandler(this.toolStripCBWorkOrderStatus_SelectedIndexChanged_1);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(188, 8);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(60, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "Customer:  ";
            // 
            // toolStripCBCustomers
            // 
            this.toolStripCBCustomers.FormattingEnabled = true;
            this.toolStripCBCustomers.Location = new System.Drawing.Point(254, 5);
            this.toolStripCBCustomers.Name = "toolStripCBCustomers";
            this.toolStripCBCustomers.Size = new System.Drawing.Size(121, 21);
            this.toolStripCBCustomers.TabIndex = 11;
            this.toolStripCBCustomers.SelectedIndexChanged += new System.EventHandler(this.toolStripCBCustomers_SelectedIndexChanged_1);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(381, 8);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(82, 13);
            this.label5.TabIndex = 12;
            this.label5.Text = "Serial Number:  ";
            // 
            // toolStripCBSerialNums
            // 
            this.toolStripCBSerialNums.FormattingEnabled = true;
            this.toolStripCBSerialNums.Location = new System.Drawing.Point(469, 5);
            this.toolStripCBSerialNums.Name = "toolStripCBSerialNums";
            this.toolStripCBSerialNums.Size = new System.Drawing.Size(121, 21);
            this.toolStripCBSerialNums.TabIndex = 13;
            this.toolStripCBSerialNums.SelectedIndexChanged += new System.EventHandler(this.toolStripCBSerialNums_SelectedIndexChanged_1);
            // 
            // WorkOrderReps
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(927, 710);
            this.Controls.Add(this.toolStripCBSerialNums);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.toolStripCBCustomers);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.toolStripCBWorkOrderStatus);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.dateTimePicker2);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.reportViewer1);
            this.DoubleBuffered = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "WorkOrderReps";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Work Order Reports";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmVEWorkOrders_FormClosing);
            this.Load += new System.EventHandler(this.frmVEWorkOrders_Load);
            this.Shown += new System.EventHandler(this.frmVEWorkOrders_Shown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dateTimePicker2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox toolStripCBWorkOrderStatus;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox toolStripCBCustomers;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox toolStripCBSerialNums;
    }
}
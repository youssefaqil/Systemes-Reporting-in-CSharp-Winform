namespace RepHp
{
    partial class Administration
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Administration));
            this.label2 = new System.Windows.Forms.Label();
            this.richTextBox3 = new System.Windows.Forms.RichTextBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.xFilesBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.reportingHpStorageFilesDataSetBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this._Reporting_Hp_StorageFilesDataSet = new RepHp._Reporting_Hp_StorageFilesDataSet();
            this.xFilesTableAdapter = new RepHp._Reporting_Hp_StorageFilesDataSetTableAdapters.XFilesTableAdapter();
            this.label1 = new System.Windows.Forms.Label();
            this.reflectionLabel1 = new DevComponents.DotNetBar.Controls.ReflectionLabel();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.buttonX2 = new DevComponents.DotNetBar.ButtonX();
            ((System.ComponentModel.ISupportInitialize)(this.xFilesBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.reportingHpStorageFilesDataSetBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this._Reporting_Hp_StorageFilesDataSet)).BeginInit();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Red;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.SystemColors.Info;
            this.label2.Location = new System.Drawing.Point(12, 93);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(233, 13);
            this.label2.TabIndex = 19;
            this.label2.Text = " Entrez le numéro de l\'état à supprimer :";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // richTextBox3
            // 
            this.richTextBox3.Location = new System.Drawing.Point(276, 188);
            this.richTextBox3.Name = "richTextBox3";
            this.richTextBox3.Size = new System.Drawing.Size(439, 85);
            this.richTextBox3.TabIndex = 34;
            this.richTextBox3.Text = "";
            // 
            // comboBox1
            // 
            this.comboBox1.DataSource = this.xFilesBindingSource;
            this.comboBox1.DisplayMember = "Number File";
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(276, 90);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(179, 21);
            this.comboBox1.TabIndex = 36;
            this.comboBox1.ValueMember = "Number File";
            // 
            // xFilesBindingSource
            // 
            this.xFilesBindingSource.DataMember = "XFiles";
            this.xFilesBindingSource.DataSource = this.reportingHpStorageFilesDataSetBindingSource;
            // 
            // reportingHpStorageFilesDataSetBindingSource
            // 
            this.reportingHpStorageFilesDataSetBindingSource.DataSource = this._Reporting_Hp_StorageFilesDataSet;
            this.reportingHpStorageFilesDataSetBindingSource.Position = 0;
            // 
            // _Reporting_Hp_StorageFilesDataSet
            // 
            this._Reporting_Hp_StorageFilesDataSet.DataSetName = "_Reporting_Hp_StorageFilesDataSet";
            this._Reporting_Hp_StorageFilesDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // xFilesTableAdapter
            // 
            this.xFilesTableAdapter.ClearBeforeFill = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Red;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.SystemColors.Info;
            this.label1.Location = new System.Drawing.Point(12, 191);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(177, 13);
            this.label1.TabIndex = 37;
            this.label1.Text = " Etat de la procédure lancée :";
            // 
            // reflectionLabel1
            // 
            this.reflectionLabel1.BackColor = System.Drawing.Color.Wheat;
            // 
            // 
            // 
            this.reflectionLabel1.BackgroundStyle.BorderBottom = DevComponents.DotNetBar.eStyleBorderType.Double;
            this.reflectionLabel1.BackgroundStyle.BorderColor = System.Drawing.SystemColors.Desktop;
            this.reflectionLabel1.BackgroundStyle.BorderLeft = DevComponents.DotNetBar.eStyleBorderType.Double;
            this.reflectionLabel1.BackgroundStyle.BorderRight = DevComponents.DotNetBar.eStyleBorderType.Double;
            this.reflectionLabel1.BackgroundStyle.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Double;
            this.reflectionLabel1.BackgroundStyle.Class = "";
            this.reflectionLabel1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.reflectionLabel1.Location = new System.Drawing.Point(110, 12);
            this.reflectionLabel1.Name = "reflectionLabel1";
            this.reflectionLabel1.Size = new System.Drawing.Size(470, 40);
            this.reflectionLabel1.TabIndex = 42;
            this.reflectionLabel1.Text = "<b><font size=\"+6\"><i>     Tableau de bord d\'administration des reporting</i><fon" +
                "t color=\"#B02B2C\"> HP</font></font></b>";
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Items.AddRange(new object[] {
            "Supprimer : Tables SQL + Fichier HP",
            "Supprimer : Tables SQL",
            "Supprimer :  Fichier HP"});
            this.comboBox2.Location = new System.Drawing.Point(489, 90);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(226, 21);
            this.comboBox2.TabIndex = 44;
            // 
            // buttonX2
            // 
            this.buttonX2.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.buttonX2.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.buttonX2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonX2.Location = new System.Drawing.Point(276, 140);
            this.buttonX2.Name = "buttonX2";
            this.buttonX2.Size = new System.Drawing.Size(179, 23);
            this.buttonX2.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.buttonX2.TabIndex = 45;
            this.buttonX2.Text = "Lancer procédure";
            this.buttonX2.Click += new System.EventHandler(this.buttonX2_Click);
            // 
            // Administration
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Azure;
            this.ClientSize = new System.Drawing.Size(729, 289);
            this.Controls.Add(this.buttonX2);
            this.Controls.Add(this.comboBox2);
            this.Controls.Add(this.reflectionLabel1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.richTextBox3);
            this.Controls.Add(this.label2);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Administration";
            this.Text = "Administration";
            this.Load += new System.EventHandler(this.Administration_Load);
            ((System.ComponentModel.ISupportInitialize)(this.xFilesBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.reportingHpStorageFilesDataSetBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this._Reporting_Hp_StorageFilesDataSet)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.RichTextBox richTextBox3;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.BindingSource reportingHpStorageFilesDataSetBindingSource;
        private _Reporting_Hp_StorageFilesDataSet _Reporting_Hp_StorageFilesDataSet;
        private System.Windows.Forms.BindingSource xFilesBindingSource;
        private _Reporting_Hp_StorageFilesDataSetTableAdapters.XFilesTableAdapter xFilesTableAdapter;
        private System.Windows.Forms.Label label1;
        private DevComponents.DotNetBar.Controls.ReflectionLabel reflectionLabel1;
        private System.Windows.Forms.ComboBox comboBox2;
        private DevComponents.DotNetBar.ButtonX buttonX2;
    }
}
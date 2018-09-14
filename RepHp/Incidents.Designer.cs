namespace RepHp
{
    partial class Incidents
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Incidents));
            this.reflectionLabel1 = new DevComponents.DotNetBar.Controls.ReflectionLabel();
            this.dataGridView1 = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.numeroEtatHPDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.incidentDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.incidentsHPBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this._Reporting_Hp_StorageFilesDataSet5 = new RepHp._Reporting_Hp_StorageFilesDataSet5();
            this.incidents_HPTableAdapter = new RepHp._Reporting_Hp_StorageFilesDataSet5TableAdapters.Incidents_HPTableAdapter();
            this.buttonX1 = new DevComponents.DotNetBar.ButtonX();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.incidentsHPBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this._Reporting_Hp_StorageFilesDataSet5)).BeginInit();
            this.SuspendLayout();
            // 
            // reflectionLabel1
            // 
            this.reflectionLabel1.BackColor = System.Drawing.Color.Beige;
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
            this.reflectionLabel1.Location = new System.Drawing.Point(56, 12);
            this.reflectionLabel1.Name = "reflectionLabel1";
            this.reflectionLabel1.Size = new System.Drawing.Size(400, 40);
            this.reflectionLabel1.TabIndex = 43;
            this.reflectionLabel1.Text = "<b><font size=\"+6\"><i>     Tableau de bord des incidents de  reporting</i><font c" +
                "olor=\"#B02B2C\"> HP</font></font></b>";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.numeroEtatHPDataGridViewTextBoxColumn,
            this.incidentDataGridViewTextBoxColumn});
            this.dataGridView1.DataSource = this.incidentsHPBindingSource;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView1.DefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(170)))), ((int)(((byte)(170)))), ((int)(((byte)(170)))));
            this.dataGridView1.Location = new System.Drawing.Point(22, 71);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(475, 181);
            this.dataGridView1.TabIndex = 44;
            // 
            // numeroEtatHPDataGridViewTextBoxColumn
            // 
            this.numeroEtatHPDataGridViewTextBoxColumn.DataPropertyName = "Numero Etat HP";
            this.numeroEtatHPDataGridViewTextBoxColumn.HeaderText = "Numero Etat HP";
            this.numeroEtatHPDataGridViewTextBoxColumn.Name = "numeroEtatHPDataGridViewTextBoxColumn";
            // 
            // incidentDataGridViewTextBoxColumn
            // 
            this.incidentDataGridViewTextBoxColumn.DataPropertyName = "Incident";
            this.incidentDataGridViewTextBoxColumn.HeaderText = "Incident";
            this.incidentDataGridViewTextBoxColumn.Name = "incidentDataGridViewTextBoxColumn";
            // 
            // incidentsHPBindingSource
            // 
            this.incidentsHPBindingSource.DataMember = "Incidents_HP";
            this.incidentsHPBindingSource.DataSource = this._Reporting_Hp_StorageFilesDataSet5;
            // 
            // _Reporting_Hp_StorageFilesDataSet5
            // 
            this._Reporting_Hp_StorageFilesDataSet5.DataSetName = "_Reporting_Hp_StorageFilesDataSet5";
            this._Reporting_Hp_StorageFilesDataSet5.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // incidents_HPTableAdapter
            // 
            this.incidents_HPTableAdapter.ClearBeforeFill = true;
            // 
            // buttonX1
            // 
            this.buttonX1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.buttonX1.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.buttonX1.Location = new System.Drawing.Point(190, 258);
            this.buttonX1.Name = "buttonX1";
            this.buttonX1.Size = new System.Drawing.Size(149, 23);
            this.buttonX1.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.buttonX1.TabIndex = 45;
            this.buttonX1.Text = "Exporter vers Excel";
            this.buttonX1.Click += new System.EventHandler(this.buttonX1_Click);
            // 
            // Incidents
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(519, 293);
            this.Controls.Add(this.buttonX1);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.reflectionLabel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Incidents";
            this.Text = "Incidents";
            this.Load += new System.EventHandler(this.Incidents_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.incidentsHPBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this._Reporting_Hp_StorageFilesDataSet5)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.Controls.ReflectionLabel reflectionLabel1;
        private DevComponents.DotNetBar.Controls.DataGridViewX dataGridView1;
        private _Reporting_Hp_StorageFilesDataSet5 _Reporting_Hp_StorageFilesDataSet5;
        private System.Windows.Forms.BindingSource incidentsHPBindingSource;
        private _Reporting_Hp_StorageFilesDataSet5TableAdapters.Incidents_HPTableAdapter incidents_HPTableAdapter;
        private System.Windows.Forms.DataGridViewTextBoxColumn numeroEtatHPDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn incidentDataGridViewTextBoxColumn;
        private DevComponents.DotNetBar.ButtonX buttonX1;
    }
}
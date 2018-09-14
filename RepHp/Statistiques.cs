using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Configuration;
using DevComponents.DotNetBar;

namespace RepHp
{
    public partial class Statistiques : Office2007Form
    {
        public Statistiques()
        {
            InitializeComponent();
        }

        private void Statistiques_Load(object sender, EventArgs e)
        {
            // TODO: cette ligne de code charge les données dans la table '_Reporting_Hp_StorageFilesDataSet4.Stat_HP'. Vous pouvez la déplacer ou la supprimer selon vos besoins.
            this.stat_HPTableAdapter1.Fill(this._Reporting_Hp_StorageFilesDataSet4.Stat_HP);
            // TODO: cette ligne de code charge les données dans la table '_Reporting_Hp_StorageFilesDataSet3.Stat_HP'. Vous pouvez la déplacer ou la supprimer selon vos besoins.
            this.stat_HPTableAdapter.Fill(this._Reporting_Hp_StorageFilesDataSet3.Stat_HP);
            // TODO: cette ligne de code charge les données dans la table '_Reporting_Hp_StorageFilesDataSet2.Statistiques'. Vous pouvez la déplacer ou la supprimer selon vos besoins.
            this.statistiquesTableAdapter.Fill(this._Reporting_Hp_StorageFilesDataSet2.Statistiques);
            SqlConnection con_Stat = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
            SqlCommand CMD_Stat = new SqlCommand("select * from [Statistiques] ", con_Stat);
            SqlDataReader Reader_Stat;
            try
            {
                con_Stat.Open();
                Reader_Stat = CMD_Stat.ExecuteReader();
                while (Reader_Stat.Read())
                {
                    this.chart1.Series["Temps_Excution_Hours"].Points.AddXY(Reader_Stat.GetInt32(0), Reader_Stat.GetInt32(1));
                    this.chart1.Series["Temps_Execution_Minutes"].Points.AddXY(Reader_Stat.GetInt32(0), Reader_Stat.GetInt32(2));

                    this.chart2.Series["Temps_Excution_Hours"].Points.AddXY(Reader_Stat.GetInt32(0) , Reader_Stat.GetInt32(1));
                    this.chart2.Series["Temps_Execution_Minutes"].Points.AddXY(Reader_Stat.GetInt32(0), Reader_Stat.GetInt32(2));

                    this.chart3.Series["Temps_Execution_Min"].Points.AddXY(Reader_Stat.GetInt32(0), Reader_Stat.GetInt32(2));
                    this.chart3.Series["Temps_Execution_Sec"].Points.AddXY(Reader_Stat.GetInt32(0), Reader_Stat.GetInt32(3));

                }

                con_Stat.Close();
            }
            catch { }
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {
        }

        private void fillByToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.statistiquesTableAdapter.FillBy(this._Reporting_Hp_StorageFilesDataSet2.Statistiques);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void Grille_Click(object sender, EventArgs e)
        {

        }
        private void ToCsV(DataGridView dGV, string filename)
        {
            string stOutput = "";
            //Export titles :
            string sHeaders = "";
            for (int j = 0; j < dGV.Columns.Count; j++)
            {
                sHeaders = sHeaders.ToString() + Convert.ToString(dGV.Columns[j].HeaderText) + "\t";
            }
            stOutput += sHeaders + "\r\n";
            // Export data.
            for (int i = 0; i < dGV.RowCount - 1; i++)
            {
                string stline = "";
                for (int j = 0; j < dGV.Rows[i].Cells.Count; j++)
                {
                    stline = stline.ToString() + Convert.ToString(dGV.Rows[i].Cells[j].Value) + "\t";
                }
                stOutput += stline + "\r\n";
            }
            Encoding utf16 = Encoding.GetEncoding(1254);
            byte[] output = utf16.GetBytes(stOutput);
            FileStream fs = new FileStream(filename, FileMode.Create);
            BinaryWriter bw = new BinaryWriter(fs);
            bw.Write(output, 0, output.Length); // write the encoded file
            bw.Flush();
            bw.Close();
            fs.Close();
        }

        private void fillBy12ToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.statistiquesTableAdapter.FillBy12(this._Reporting_Hp_StorageFilesDataSet2.Statistiques);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void chart2_Click(object sender, EventArgs e)
        {

        }

      

        private void button2_Click(object sender, EventArgs e)
        {
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
           
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls) |*.xls";
            sfd.FileName = "Duration_Execution_HP_Reports.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                ToCsV(dataGridView1, sfd.FileName);
                MessageBox.Show("Le fichier " + sfd.FileName + " a été créé avec succès !");
            }
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

      
        
      
    }
}

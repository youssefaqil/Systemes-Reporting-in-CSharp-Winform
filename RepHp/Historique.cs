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
    public partial class Historique : Office2007Form
    {
        public Historique()
        {
            InitializeComponent();
        }

        private void Historique_Load(object sender, EventArgs e)
        {
            DataGrid_Fill();
        }
        SqlDataAdapter sda_SIT;
        DataSet ds_SIT;
        SqlCommandBuilder cmd_SIT;
        SqlDataAdapter sda_AMS;
        DataSet ds_AMS;
        SqlCommandBuilder cmd_AMS;
        SqlDataAdapter sda_CUS;
        DataSet ds_CUS;
        //SqlCommandBuilder cmd_CUS;
        SqlDataAdapter sda_PRS;
        DataSet ds_PRS;
        //SqlCommandBuilder cmd_PRS;
        SqlDataAdapter sda_HDR;
        DataSet ds_HDR;
        //SqlCommandBuilder cmd_HDR;
        SqlDataAdapter sda_TRL;
        DataSet ds_TRL;
        SqlCommandBuilder cmd_TRL;
        /* Fonction de remplissage de la grille avec l'historique des fichiers dèja crées */
        private void DataGrid_Fill()

        {
            SqlConnection con_dg = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
            con_dg.Open();
            SqlDataAdapter sda = new SqlDataAdapter("SELECT [Number File] ,[Date création] ,[Date début] ,[Date fin],[Type fichier] FROM [Reporting-Hp-StorageFiles].[dbo].[XFiles] order by 1 desc", con_dg);
            DataTable dt = new DataTable();
            sda.Fill(dt);
            dataGridView1.DataSource = dt;
            con_dg.Close();
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

        private void button1_Click(object sender, EventArgs e)
        {
            DataGrid_Fill();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls) |*.xls";
            sfd.FileName = "Historique_HP_Reports.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                ToCsV(dataGridView1, sfd.FileName);
                MessageBox.Show("Le fichier " + sfd.FileName + " a été créé avec succès !");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                /*Chargement de la table Header*/
                SqlConnection con_dg = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                sda_HDR = new SqlDataAdapter("SELECT * from [" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "HDR]", con_dg);
                ds_HDR = new System.Data.DataSet();
                sda_HDR.Fill(ds_HDR, "[" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "HDR]");
                dataGridView2.DataSource = ds_HDR.Tables[0];
              
                /* Chargement de la table SIT */
                sda_SIT = new SqlDataAdapter("SELECT * from [" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "SIT]", con_dg);
                ds_SIT = new System.Data.DataSet();
                sda_SIT.Fill(ds_SIT, "[" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "SIT]");
                dataGridView3.DataSource = ds_SIT.Tables[0];
                


                /* Chargement de la table AMS */
                sda_AMS = new SqlDataAdapter("SELECT * from [" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "AMS]", con_dg);
                ds_AMS = new System.Data.DataSet();
                sda_AMS.Fill(ds_AMS, "[" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "AMS]");
                dataGridView4.DataSource = ds_AMS.Tables[0];
                

                /* Chargement de la table PRS */
                sda_PRS = new SqlDataAdapter("SELECT * from [" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "PRS]", con_dg);
                ds_PRS = new System.Data.DataSet();
                sda_PRS.Fill(ds_PRS, "[" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "PRS]");
                dataGridView5.DataSource = ds_PRS.Tables[0];

                /* Chargement de la table CUS */
                sda_CUS = new SqlDataAdapter("SELECT * from [" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "CUS]", con_dg);
                ds_CUS = new System.Data.DataSet();
                sda_CUS.Fill(ds_CUS, "[" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "CUS]");
                dataGridView6.DataSource = ds_CUS.Tables[0];
                

                /* Chargement de la table TRL */
                sda_TRL = new SqlDataAdapter("SELECT * from [" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "TRL]", con_dg);
                ds_TRL = new System.Data.DataSet();
                sda_TRL.Fill(ds_TRL, "[" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "TRL]");
                dataGridView7.DataSource = ds_TRL.Tables[0];
                con_dg.Close();
            }
            catch
            {
                MessageBox.Show("Le fichier selectionné n'existe pas sur la base !");
                dataGridView2.DataSource = null;
                dataGridView3.DataSource = null;
                dataGridView3.DataSource = null;
                dataGridView4.DataSource = null;
                dataGridView5.DataSource = null;
                dataGridView6.DataSource = null;
            }
        }

      

        private void button11_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Text Documents (*.txt) |*.txt";
            ofd.FileName = @"C:\Users\bbourquia\Desktop\Reports_HP\iflash2_matel." +  dataGridView1.CurrentRow.Cells["Number File"].Value.ToString()  + ".txt";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
               /* File.Open(@"C:\Users\bbourquia\Desktop\Reports_HP\iflash2_matel." +  dataGridView1.CurrentRow.Cells["Number File"].Value.ToString()  + ".txt");*/
               
            }

        }

        private void buttonX1_Click_1(object sender, EventArgs e)
        {
            try
            {

                cmd_SIT = new SqlCommandBuilder(sda_SIT);
                sda_SIT.Update(ds_SIT, "[" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "SIT]");
                MessageBox.Show("Ligne modifiée dans la table :" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "SIT");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void buttonX2_Click_1(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls) |*.xls";
            sfd.FileName = "" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "_Report_HP_SIT.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                ToCsV(dataGridView3, sfd.FileName);
                MessageBox.Show("Le fichier " + sfd.FileName + " a été créé avec succès !");
            }
        }

        private void buttonX3_Click_1(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls) |*.xls";
            sfd.FileName = "" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "_Report_HP_HDR.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                ToCsV(dataGridView2, sfd.FileName);
                MessageBox.Show("Le fichier " + sfd.FileName + " a été créé avec succès !");
            }

        }

        private void buttonX5_Click_1(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls) |*.xls";
            sfd.FileName = "" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "_Report_HP_PRS.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                ToCsV(dataGridView5, sfd.FileName);
                MessageBox.Show("Le fichier " + sfd.FileName + " a été créé avec succès !");
            }
        }

        private void buttonX4_Click_1(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls) |*.xls";
            sfd.FileName = "" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "_Report_HP_AMS.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                ToCsV(dataGridView4, sfd.FileName);
                MessageBox.Show("Le fichier " + sfd.FileName + " a été créé avec succès !");
            }
        }

        private void buttonX6_Click_1(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls) |*.xls";
            sfd.FileName = "" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "_Report_HP_CUS.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                ToCsV(dataGridView6, sfd.FileName);
                MessageBox.Show("Le fichier " + sfd.FileName + " a été créé avec succès !");
            }
        }

        private void buttonX7_Click_1(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls) |*.xls";
            sfd.FileName = "" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "_Report_HP_TRL.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                ToCsV(dataGridView7, sfd.FileName);
                MessageBox.Show("Le fichier " + sfd.FileName + " a été créé avec succès !");
            }
        }

        private void buttonX8_Click(object sender, EventArgs e)
        {
            try
            {

                cmd_AMS = new SqlCommandBuilder(sda_AMS);
                sda_AMS.Update(ds_AMS, "[" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "AMS]");
                MessageBox.Show("Ligne modifiée dans la table :" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "AMS");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void buttonX9_Click(object sender, EventArgs e)
        {
            try
            {

                cmd_TRL = new SqlCommandBuilder(sda_TRL);
                sda_TRL.Update(ds_TRL, "[" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "TRL]");
                MessageBox.Show("Ligne modifiée dans la table :" + dataGridView1.CurrentRow.Cells["Number File"].Value.ToString() + "TRL");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

     
    }
}

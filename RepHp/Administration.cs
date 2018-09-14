using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using DevComponents.DotNetBar;

namespace RepHp
{
    public partial class Administration : Office2007Form
    {
        public Administration()
        {
            InitializeComponent();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void Administration_Load(object sender, EventArgs e)
        {
            // TODO: cette ligne de code charge les données dans la table '_Reporting_Hp_StorageFilesDataSet.XFiles'. Vous pouvez la déplacer ou la supprimer selon vos besoins.
            this.xFilesTableAdapter.Fill(this._Reporting_Hp_StorageFilesDataSet.XFiles);

        }

        private void button1_Click(object sender, EventArgs e)
        {
}

      

        private void buttonX2_Click(object sender, EventArgs e)
        {
            int h = comboBox2.SelectedIndex;
            switch(h)
            {
                case 0 :
                    MessageBox.Show("Lancer Suppression : Tables SQL + Fichier HP ?");

                    if (System.IO.File.Exists(@"\\wayvs\Reporting\\iflashas2_matel." + comboBox1.Text + ".dat"))
                    {
                        // Suupression des tables Sql 
                        SqlConnection Con_rmp_TRL = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                        Con_rmp_TRL.Open();
                        if (comboBox1.Text.Contains("HPE"))
                        {
                            SqlCommand CMD_HDR_E2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPE", "HDR_E]"), Con_rmp_TRL);
                            SqlCommand CMD_CUS_E2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPE", "CUS_E]"), Con_rmp_TRL);
                            SqlCommand CMD_SIT_E2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPE", "SIT_E]"), Con_rmp_TRL);
                            SqlCommand CMD_PRS_E2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPE", "PRS_E]"), Con_rmp_TRL);
                            SqlCommand CMD_SNT_E2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPE", "SNT_E]"), Con_rmp_TRL);
                            SqlCommand CMD_TRL_E2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPE", "TRL_E]"), Con_rmp_TRL);
                            SqlCommand CMD_AMS_E2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPE", "AMS_E]"), Con_rmp_TRL);
                            CMD_HDR_E2.ExecuteNonQuery();
                            richTextBox3.Text = "Table [" + comboBox1.Text.Replace("_HPI", "HDR_E]") + " supprimée";

                            CMD_CUS_E2.ExecuteNonQuery();
                            richTextBox3.Text += " / Table [" + comboBox1.Text.Replace("_HPI", "CUS_E]") + " supprimée";

                            CMD_SIT_E2.ExecuteNonQuery();
                            richTextBox3.Text += " / Table [" + comboBox1.Text.Replace("_HPI", "SIT_E]") + " supprimée";

                            CMD_PRS_E2.ExecuteNonQuery();
                            richTextBox3.Text += " / Table [" + comboBox1.Text.Replace("_HPI", "PRS_E]") + " supprimée";

                            CMD_SNT_E2.ExecuteNonQuery();
                            richTextBox3.Text += " / Table [" + comboBox1.Text.Replace("_HPI", "SNT_E]") + " supprimée";

                            CMD_TRL_E2.ExecuteNonQuery();
                            richTextBox3.Text += " / Table [" + comboBox1.Text.Replace("_HPI", "TRL_E]") + " supprimée";

                            CMD_AMS_E2.ExecuteNonQuery();
                            richTextBox3.Text += " / Table [" + comboBox1.Text.Replace("_HPE", "AMS_E]") + " supprimée";
                        }
                        else
                        {
                            SqlCommand CMD_HDR_I2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPI", "HDR_I]"), Con_rmp_TRL);
                            SqlCommand CMD_CUS_I2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPI", "CUS_I]"), Con_rmp_TRL);
                            SqlCommand CMD_SIT_I2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPI", "SIT_I]"), Con_rmp_TRL);
                            SqlCommand CMD_PRS_I2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPI", "PRS_I]"), Con_rmp_TRL);
                            SqlCommand CMD_SNT_I2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPI", "SNT_I]"), Con_rmp_TRL);
                            SqlCommand CMD_TRL_I2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPI", "TRL_I]"), Con_rmp_TRL);
                            SqlCommand CMD_AMS_I2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPI", "AMS_I]"), Con_rmp_TRL);

                            CMD_HDR_I2.ExecuteNonQuery();
                            richTextBox3.Text = "Table [" + comboBox1.Text.Replace("_HPI", "AMS_I]") + " supprimée";

                            CMD_CUS_I2.ExecuteNonQuery();
                            richTextBox3.Text += " / Table [" + comboBox1.Text.Replace("_HPI", "CUS_I]") + " supprimée";

                            CMD_SIT_I2.ExecuteNonQuery();
                            richTextBox3.Text += " / Table [" + comboBox1.Text.Replace("_HPI", "SIT_I]") + " supprimées";

                            CMD_PRS_I2.ExecuteNonQuery();
                            richTextBox3.Text += " / Table [" + comboBox1.Text.Replace("_HPI", "PRS_I]") + " supprimée";

                            CMD_SNT_I2.ExecuteNonQuery();
                            richTextBox3.Text += " / Table [" + comboBox1.Text.Replace("_HPI", "SNT_I]") + " supprimée";

                            CMD_TRL_I2.ExecuteNonQuery();
                            richTextBox3.Text += " / Table [" + comboBox1.Text.Replace("_HPI", "TRL_I]") + " supprimée";

                            CMD_AMS_I2.ExecuteNonQuery();
                            richTextBox3.Text += " / Table [" + comboBox1.Text.Replace("_HPI", "AMS_I]") + " supprimée";
                        }
                        SqlCommand CMD_Xfiles = new SqlCommand("DELETE FROM XFiles WHERE [Number file]='" + comboBox1.Text + "'", Con_rmp_TRL);

                        CMD_Xfiles.ExecuteNonQuery();
                        richTextBox3.Text += "La ligne [" + comboBox1.Text + "] a été bien supprimée de [Xfiles]";
                        Con_rmp_TRL.Close();
                        // Suupression du fichier HP
                        System.IO.File.Delete((@"\\wayvs\Reporting\\iflashas2_matel." + comboBox1.Text + ".dat"));
                        richTextBox3.Text += "   Etat numéro < " + comboBox1.Text + " > supprimé avec succès ! ";
                    }
                    else
                    {
                        richTextBox3.Clear();
                        richTextBox3.Text = "Etat numéro < " + comboBox1.Text + " > n'as pas été supprimé car il n'existe pas ou a dèjà été supprimé !";

                    }
                    break;


                case 1 :
                    MessageBox.Show("Lancer suppression des tables numéro : " + comboBox1.Text + " du report HP ?");
                    // Suupression des tables Sql 
                    SqlConnection Con_rmp_TRL2 = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                    Con_rmp_TRL2.Open();
                    if (comboBox1.Text.Contains("HPE"))
                    {
                        SqlCommand CMD_HDR_E2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPE", "HDR_E]"), Con_rmp_TRL2);
                        SqlCommand CMD_CUS_E2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPE", "CUS_E]"), Con_rmp_TRL2);
                        SqlCommand CMD_SIT_E2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPE", "SIT_E]"), Con_rmp_TRL2);
                        SqlCommand CMD_PRS_E2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPE", "PRS_E]"), Con_rmp_TRL2);
                        SqlCommand CMD_SNT_E2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPE", "SNT_E]"), Con_rmp_TRL2);
                        SqlCommand CMD_TRL_E2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPE", "TRL_E]"), Con_rmp_TRL2);
                        SqlCommand CMD_AMS_E2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPE", "AMS_E]"), Con_rmp_TRL2);
                        
                        CMD_HDR_E2.ExecuteNonQuery();
                        richTextBox3.Text = "Table [" + comboBox1.Text.Replace("_HPI", "HDR_E]") + " supprimée";

                        CMD_CUS_E2.ExecuteNonQuery();
                        richTextBox3.Text += " / Table [" + comboBox1.Text.Replace("_HPI", "CUS_E]") + " supprimée";

                        CMD_SIT_E2.ExecuteNonQuery();
                        richTextBox3.Text += " / Table [" + comboBox1.Text.Replace("_HPI", "SIT_E]") + " supprimée";

                        CMD_PRS_E2.ExecuteNonQuery();
                        richTextBox3.Text += " / Table [" + comboBox1.Text.Replace("_HPI", "PRS_E]") + " supprimée";

                        CMD_SNT_E2.ExecuteNonQuery();
                        richTextBox3.Text += " / Table [" + comboBox1.Text.Replace("_HPI", "SNT_E]") + " supprimée";

                        CMD_TRL_E2.ExecuteNonQuery();
                        richTextBox3.Text += " / Table [" + comboBox1.Text.Replace("_HPI", "TRL_E]") + " supprimée";

                        CMD_AMS_E2.ExecuteNonQuery();
                        richTextBox3.Text += " / Table [" + comboBox1.Text.Replace("_HPE", "AMS_E]") + " supprimée";
                    }
                    else
                    {
                        SqlCommand CMD_HDR_I2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPI", "HDR_I]"), Con_rmp_TRL2);
                        SqlCommand CMD_CUS_I2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPI", "CUS_I]"), Con_rmp_TRL2);
                        SqlCommand CMD_SIT_I2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPI", "SIT_I]"), Con_rmp_TRL2);
                        SqlCommand CMD_PRS_I2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPI", "PRS_I]"), Con_rmp_TRL2);
                        SqlCommand CMD_SNT_I2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPI", "SNT_I]"), Con_rmp_TRL2);
                        SqlCommand CMD_TRL_I2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPI", "TRL_I]"), Con_rmp_TRL2);
                        SqlCommand CMD_AMS_I2 = new SqlCommand("DROP TABLE [" + comboBox1.Text.Replace("_HPI", "AMS_I]"), Con_rmp_TRL2);

                        CMD_HDR_I2.ExecuteNonQuery();
                        richTextBox3.Text = "Table [" + comboBox1.Text.Replace("_HPI", "AMS_I]") + " supprimée";

                        CMD_CUS_I2.ExecuteNonQuery();
                        richTextBox3.Text += " / Table [" + comboBox1.Text.Replace("_HPI", "CUS_I]") + " supprimée";

                        CMD_SIT_I2.ExecuteNonQuery();
                        richTextBox3.Text += " / Table [" + comboBox1.Text.Replace("_HPI", "SIT_I]") + " supprimées";

                        CMD_PRS_I2.ExecuteNonQuery();
                        richTextBox3.Text += " / Table [" +comboBox1.Text.Replace("_HPI", "PRS_I]") + " supprimée";

                        CMD_SNT_I2.ExecuteNonQuery();
                        richTextBox3.Text += " / Table [" + comboBox1.Text.Replace("_HPI", "SNT_I]") + " supprimée";

                        CMD_TRL_I2.ExecuteNonQuery();
                        richTextBox3.Text += " / Table [" + comboBox1.Text.Replace("_HPI", "TRL_I]") + " supprimée";

                        CMD_AMS_I2.ExecuteNonQuery();
                        richTextBox3.Text += " / Table [" + comboBox1.Text.Replace("_HPI", "AMS_I]") + " supprimée";
                    }
                    SqlCommand CMD_Xfiles2 = new SqlCommand("DELETE FROM XFiles WHERE [Number file]='" + comboBox1.Text + "'", Con_rmp_TRL2);

                    CMD_Xfiles2.ExecuteNonQuery();
                    richTextBox3.Text += "La ligne [" + comboBox1.Text +"] a été bien supprimée de [Xfiles]";

                    Con_rmp_TRL2.Close();
                    break;

                case 2:
                    MessageBox.Show("Lancer suppression du Fichier HP : iflashas2_matel." + comboBox1.Text + ".dat !");
                    if (System.IO.File.Exists(@"\\wayvs\Reporting\\iflashas2_matel." + comboBox1.Text + ".dat"))
                    {
                        System.IO.File.Delete((@"\\wayvs\Reporting\\iflashas2_matel." + comboBox1.Text + ".dat"));
                        richTextBox3.Clear();
                        richTextBox3.Text = "Etat numéro < " + comboBox1.Text + " > supprimé avec succès ! ";
                    }
                    else
                    {
                        richTextBox3.Clear();
                        richTextBox3.Text = "Etat numéro < " + comboBox1.Text + " > n'as pas été supprimé car il n'existe pas ou a dèjà été supprimé !";

                    }
                    break;
                default :
                    MessageBox.Show("Veuillez sélectionner une option de suppression !");
                    break;
            }


        }
    }
        }

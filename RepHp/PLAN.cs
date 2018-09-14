using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using DevComponents.DotNetBar;
using System.IO;
using System.Data.SqlClient;
namespace RepHp
{
    public partial class PLAN : Office2007Form
    {
        public PLAN()
        {
            InitializeComponent();
        }
        /*class WednesdayTimer
        {
            private System.Timers.Timer timer;
            private DateTime dueDate;


            public WednesdayTimer(DateTime dueDate)
            {
                this.dueDate = dueDate;
                this.timer = new System.Timers.Timer(TimeSpan.FromMinutes(1).TotalMilliseconds);
                this.timer.Elapsed += this.timerElapsed;
                this.timer.AutoReset = false;
                this.timer.Enabled = true;


            }
            private void timerElapsed(object sender, System.Timers.ElapsedEventArgs e)
            {
                if (DateTime.Now < this.dueDate)
                    return;
                this.timer.Stop();
                this.timer.Elapsed -= this.timerElapsed;
                this.timer.Dispose();
            }
            public void Start()
            {
                this.timer.Start();
            }
        }
         * */
       

        private void PLAN_Load(object sender, EventArgs e)
        {
            // TODO: cette ligne de code charge les données dans la table '_Reporting_Hp_StorageFilesDataSet.XFiles'. Vous pouvez la déplacer ou la supprimer selon vos besoins.
            this.xFilesTableAdapter.Fill(this._Reporting_Hp_StorageFilesDataSet.XFiles);
           
        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            try
            {
                string text = File.ReadAllText(@"\\wayvs\Reporting\iflashas2_matel." + comboBox1.Text.ToString() + ".dat");
                richTextBox1.Text = text;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
               
            
        }
        /*Fonction de création et de remplissage du fichier */
        private bool MAJ_Report_HP()
        {
            try
            {
              
                SqlConnection Con_HP = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                string Query_HP_HDR = "select * from [" + comboBox1.Text.ToString() + "HDR] ";
                string Query_HP_CUS = "select * from [" + comboBox1.Text.ToString() + "CUS]";
                string Query_HP_SIT = "select * from [" + comboBox1.Text.ToString() + "SIT]";
                string Query_HP_AMS = "select * from [" + comboBox1.Text.ToString() + "AMS]";
                string Query_HP_PRS = "select * from [" + comboBox1.Text.ToString() + "PRS]";
                string Query_HP_TRL = "select * from [" + comboBox1.Text.ToString() + "TRL]";
                SqlCommand Cmd_HDR = new SqlCommand(Query_HP_HDR, Con_HP);
                SqlCommand Cmd_CUS = new SqlCommand(Query_HP_CUS, Con_HP);
                SqlCommand Cmd_SIT = new SqlCommand(Query_HP_SIT, Con_HP);
                SqlCommand Cmd_AMS = new SqlCommand(Query_HP_AMS, Con_HP);
                SqlCommand Cmd_PRS = new SqlCommand(Query_HP_PRS, Con_HP);
                SqlCommand Cmd_TRL = new SqlCommand(Query_HP_TRL, Con_HP);
                File.Delete(@"\\wayvs\Reporting\iflashas2_matel." + comboBox1.Text.ToString() + ".dat");
                richTextBox2.Text = "Etat HP numéro : " + comboBox1.Text.ToString() + " vidé ! ";
                FileStream TheFile = File.Create(@"\\wayvs\Reporting\iflashas2_matel." + comboBox1.Text.ToString() + ".dat");
                StreamWriter Writer = new StreamWriter(TheFile, Encoding.GetEncoding(1252));


                string Header = "HDR";
                string HDR = Header.Trim();

                using (Con_HP)
                {
                    Con_HP.Open();
                    /* Ecriture du HDR */
                    using (SqlDataReader Reader_HP_HDR = Cmd_HDR.ExecuteReader())
                    using (Writer)
                    {
                        while (Reader_HP_HDR.Read())
                        {
                            Writer.WriteLine(HDR.ToUpper().ToString() + Reader_HP_HDR[0].ToString() + "00" + Reader_HP_HDR[1].ToString() + Reader_HP_HDR[2].ToString() + Reader_HP_HDR[3].ToString() + Reader_HP_HDR[4].ToString() + Reader_HP_HDR[5].ToString() + Reader_HP_HDR[6].ToString() + Reader_HP_HDR[7].ToString() + Reader_HP_HDR[8].ToString() + Reader_HP_HDR[9].ToString() + Reader_HP_HDR[10].ToString() + Reader_HP_HDR[11].ToString());
                        }

                        /* Ecriture du CUS parès HDR*/
                        if (!Reader_HP_HDR.Read())
                        {
                            Reader_HP_HDR.Close();
                            SqlDataReader Reader_HP_SIT = Cmd_SIT.ExecuteReader();
                            while (Reader_HP_SIT.Read())
                            {     /* si le reserved inventory est supérieur au total inventory on les rends égaux */
                                if (int.Parse(Reader_HP_SIT[5].ToString()) > int.Parse(Reader_HP_SIT[4].ToString()))
                                {
                                    Writer.WriteLine("SIT" + Reader_HP_SIT[0].ToString() + Reader_HP_SIT[1].ToString() + Reader_HP_SIT[2].ToString() + Reader_HP_SIT[3].ToString() + Reader_HP_SIT[4].ToString() + Reader_HP_SIT[4].ToString() + Reader_HP_SIT[6].ToString() + Reader_HP_SIT[7].ToString() + Reader_HP_SIT[8].ToString() + Reader_HP_SIT[9].ToString());
                                }
                                else
                                {
                                    Writer.WriteLine("SIT" + Reader_HP_SIT[0].ToString() + Reader_HP_SIT[1].ToString() + Reader_HP_SIT[2].ToString() + Reader_HP_SIT[3].ToString() + Reader_HP_SIT[4].ToString() + Reader_HP_SIT[5].ToString() + Reader_HP_SIT[6].ToString() + Reader_HP_SIT[7].ToString() + Reader_HP_SIT[8].ToString() + Reader_HP_SIT[9].ToString());
                                }
                            }
                            /*  Reader_HP_CUS.Close(); */

                            /* Ecriture du AMS parès HDR*/
                            if (!Reader_HP_SIT.Read())
                            {
                                Reader_HP_SIT.Close();
                                SqlDataReader Reader_HP_CUS = Cmd_CUS.ExecuteReader();
                                while (Reader_HP_CUS.Read())
                                {
                                    Writer.WriteLine("CUS" + Reader_HP_CUS[0].ToString() + Reader_HP_CUS[1].ToString() + Reader_HP_CUS[2].ToString().ToUpper() + Reader_HP_CUS[3].ToString().ToUpper() + Reader_HP_CUS[4].ToString() + Reader_HP_CUS[5].ToString().ToUpper() + Reader_HP_CUS[6].ToString() + Reader_HP_CUS[7].ToString() + Reader_HP_CUS[8].ToString());
                                }
                                /*   Reader_HP_AMS.Close(); */
                                /* Ecriture du PRS après AMS*/
                                if (!Reader_HP_CUS.Read())
                                {
                                    Reader_HP_CUS.Close();
                                    SqlDataReader Reader_HP_PRS = Cmd_PRS.ExecuteReader();
                                    while (Reader_HP_PRS.Read())
                                    {
                                        Writer.WriteLine("PRS" + Reader_HP_PRS[0].ToString() + Reader_HP_PRS[1].ToString() + Reader_HP_PRS[2].ToString() + Reader_HP_PRS[3].ToString() + Reader_HP_PRS[4].ToString() + Reader_HP_PRS[5].ToString() + Reader_HP_PRS[6].ToString() + Reader_HP_PRS[7].ToString() + Reader_HP_PRS[8].ToString() + "00000000000000.00" + Reader_HP_PRS[10].ToString());
                                    }
                                    /*  Reader_HP_PRS.Close(); */

                                    /* Ecriture du SIT après PRS*/
                                    if (!Reader_HP_PRS.Read())
                                    {
                                        Reader_HP_PRS.Close();
                                        SqlDataReader Reader_HP_AMS = Cmd_AMS.ExecuteReader();
                                        while (Reader_HP_AMS.Read())
                                        {
                                            Writer.WriteLine("AMS" + Reader_HP_AMS[0].ToString() + Reader_HP_AMS[1].ToString() + Reader_HP_AMS[2].ToString() + Reader_HP_AMS[3].ToString());
                                        }
                                        /* Reader_HP_SIT.Close(); */

                                        if (!Reader_HP_AMS.Read())
                                        {
                                            Reader_HP_AMS.Close();
                                            SqlDataReader Reader_HP_TRL = Cmd_TRL.ExecuteReader();
                                            while (Reader_HP_TRL.Read())
                                            {
                                                Writer.WriteLine("TRL" + Reader_HP_TRL[0].ToString());
                                            }
                                            Reader_HP_TRL.Close();
                                        }
                                    }

                                }
                            }

                        }

                    }
                    richTextBox2.Text += "/ Etat HP mis à jour !";
                
                    return true;

                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Etat HP non crée !");
                MessageBox.Show("Exception : " + e.Message + " ");
                return false;
            }
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            MAJ_Report_HP();
        }
         private bool  Depos_HP()
        {
            try
            {
                File.Copy(@"\\wayvs\Reporting\iflashas2_matel." + comboBox1.Text.ToString() + ".dat", @"\\wayvs\Reporting\iflashas2_matel_bis." + comboBox1.Text.ToString() + ".dat");
                System.IO.File.Move(@"\\wayvs\Reporting\iflashas2_matel_bis." + comboBox1.Text.ToString() + ".dat", @"\\edi\application-octet-stream\iflashas2_matel." + comboBox1.Text.ToString() + ".dat");
                File.Delete(@"\\wayvs\Reporting\iflashas2_matel_bis." + comboBox1.Text.ToString() + ".dat");
                return true;
               
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return false;
            }
         }
        private void buttonX3_Click(object sender, EventArgs e)
        {
            try
            {
                Depos_HP();
                MessageBox.Show("Fichier numéro '" + comboBox1.Text.ToString() + "' correctement déposé dans EDI ");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        }
        }


      
  


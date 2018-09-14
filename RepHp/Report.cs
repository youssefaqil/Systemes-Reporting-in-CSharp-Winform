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
using System.Threading;
using Microsoft.VisualBasic;
using DevComponents.DotNetBar;
using System.Net.Mail;
using Excel = Microsoft.Office.Interop.Excel; 
namespace RepHp
{
    public partial class Report : Office2007Form
    {
        /* paramètres système de notification */
        public SmtpClient client = new SmtpClient();
        public MailMessage msg = new MailMessage();
        public System.Net.NetworkCredential smtpCreds = new System.Net.NetworkCredential("notifhp@disway.com", "");
       /* Fin paramètres  */
       
        /*
         *** ZHA 16/062015 NEW DEV HPE and HPI ***
         */
        string CodeHPE = null, CodeHPI = null;

        public void Get_typeHP()
        {
            SqlConnection Config = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
            string Query = "select Valeur_information from [Informations] where Type_information in ('CodeHPE','CodeHPI')";
            SqlCommand Cmd_config = new SqlCommand(Query, Config);
            Config.Open();
            SqlDataReader Reader_config = Cmd_config.ExecuteReader();
            while (Reader_config.Read())
            {
                if (Reader_config[0].ToString() == "HPE")  
                    CodeHPE = Reader_config[0].ToString();
                else
                    CodeHPI = Reader_config[0].ToString();
            }
            Reader_config.Close();
            Config.Close();
        }
        /* Fin DEV */

        /* Fonction d'envoi d'une notification mail */
        public void SendEmail(string sendTo, string sendFrom, string subject, string body)
        {
            try
            {
                //Setup SMTP
                client.Host = "192.168.50.62";
                client.Port = 25;
                /* client.UseDefaultCredentials = false;
                client.Credentials = smtpCreds; */
                /*client.EnableSsl = true; */

                //Setup strings to MailAdress
                MailAddress to = new MailAddress(sendTo);
                MailAddress from = new MailAddress(sendFrom);

                //Setup message settings
                msg.Subject = subject;
                msg.Body = body;
                msg.From = from;
                msg.To.Add(to);

                //Send email
                client.Send(msg);


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "ERROR");
            }

        }
        /* Fin de la fonction */

        public Report()
        {
            InitializeComponent();
            timer1.Start();
            Select_Number_File();
            String Time_Creation_File = DateTime.Now.ToString("HH:mm:ss");
            String Date_Creation_File = DateTime.Now.ToShortDateString();
            textBox1.Text = "20:00:00";
            textBox2.Text = Date_Creation_File;
        }
    

        private void Report_Load(object sender, EventArgs e)
        {
            // TODO: cette ligne de code charge les données dans la table '_Reporting_Hp_StorageFilesDataSet1.FILETYPE'. Vous pouvez la déplacer ou la supprimer selon vos besoins.
         //   this.fILETYPETableAdapter.Fill(this._Reporting_Hp_StorageFilesDataSet1.FILETYPE);
           /* String Time_Creation_File = DateTime.Now.ToString("HH:mm:ss");
            String Date_Creation_File = DateTime.Now.ToShortDateString();
            textBox1.Text = Time_Creation_File;
            textBox2.Text = Date_Creation_File; */
            Get_typeHP();
        }
        /* Fonction qui retourne le numéro de fichier du dernier report crée */
        private void Select_Number_File()
        {
            SqlConnection con2 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
            con2.Open();
            SqlCommand cmd = new SqlCommand("select Valeur_information from Informations where Type_information='LASTFILESEQUENCENUMBER'", con2);
            string NumberFile = (string)cmd.ExecuteScalar();
            con2.Close();
            textBox5.Text = NumberFile;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            DateTime dateTime = DateTime.Now;
            this.textBox4.Text = "      " + dateTime.ToString() + "";
        }

        /* Fonction d'incrémentation du numéro de fichier */
        private void Increment_File_ID()
        {
            SqlConnection con3 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
            con3.Open();
            int a = int.Parse(textBox5.Text);
            a++;
            SqlCommand Cmd_MAJ = new SqlCommand("update Informations set Valeur_information = " + a + " where Type_information='LASTFILESEQUENCENUMBER'", con3);
            Cmd_MAJ.ExecuteNonQuery();
            con3.Close();
        }

        /* Fonction qui stocke dans la bd les infos sur le fichier crée */
        private void Stockage_File(string codeHP)
        {

            SqlConnection con4 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
            con4.Open();
            if (codeHP == CodeHPE)
            {
                if (comboBox1.SelectedIndex == 0)
                {
                    SqlCommand Cmd_Store_Data_File = new SqlCommand
                    ("insert into XFiles ([Number File],[Date création], [Date début],[Date fin],[Type fichier]) values ('" + textBox5.Text + "_" + codeHP + "','" + textBox4.Text + "' ,'" + dateTimePicker1.Text + "','" + dateTimePicker2.Text + "','PRODUCTION')", con4);
                    Cmd_Store_Data_File.ExecuteNonQuery();

                }
                else
                {
                    SqlCommand Cmd_Store_Data_File2 = new SqlCommand
                                   ("insert into XFiles ([Number File], [Date création], [Date début],[Date fin],[Type fichier]) values ('" + textBox5.Text + "_" + codeHP + "','" + textBox4.Text + "' ,'" + dateTimePicker1.Text + "','" + dateTimePicker2.Text + "','TEST')", con4);
                    Cmd_Store_Data_File2.ExecuteNonQuery();

                }
            }
            else
            {
                if (comboBox1.SelectedIndex == 0)
                {
                    SqlCommand Cmd_Store_Data_File = new SqlCommand
                    ("insert into XFiles ([Number File], [Date création], [Date début],[Date fin],[Type fichier]) values ('" + textBox5.Text + "_" + codeHP + "','" + textBox4.Text + "' ,'" + dateTimePicker1.Text + "','" + dateTimePicker2.Text + "','PRODUCTION')", con4);
                    Cmd_Store_Data_File.ExecuteNonQuery();

                }
                else
                {
                    SqlCommand Cmd_Store_Data_File2 = new SqlCommand
                                   ("insert into XFiles ([Number File], [Date création], [Date début],[Date fin],[Type fichier]) values ('" + textBox5.Text + "_" + codeHP + "','" + textBox4.Text + "' ,'" + dateTimePicker1.Text + "','" + dateTimePicker2.Text + "','TEST')", con4);
                    Cmd_Store_Data_File2.ExecuteNonQuery();

                }
            }
            con4.Close();
        }

        /* Fonction de création de la table HDR_E */
        private bool Creation_HDR_Table(string codeHP)
        {
            try
            {
                if (codeHP == CodeHPE)
                {
                    SqlConnection con5 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                    con5.Open();
                    string myquery = "CREATE TABLE [" + textBox5.Text + "HDR_E] (FileLayoutID Varchar(50) ,FileLayoutVersion   Varchar (50) ,PartnerId   Varchar (50) ,PartnerName   Varchar (50) ,Partnercontact   Varchar (50) ,PartnerReferenceDepotID   Varchar (50) ,FileCreationDate   Varchar (50) ,FileCreationTime  Varchar (50) ,PeriodStartDate   Varchar (50) ,PeriodEndDate   Varchar (50) ,FileSequenceNumber   Varchar (50) PRIMARY KEY ,TestFileIndicator   Varchar (50))";
                    SqlCommand Cmd_HDR = new SqlCommand(myquery, con5);
                    Cmd_HDR.ExecuteNonQuery();
                    con5.Close();
                    /* MessageBox.Show("OK HDR_E"); */
                    richTextBox3.Text = "<Création> : " + "OK HDR_E";
                    SendEmail(" notifhp@disway.com", "notifhp@disway.com", "Alerte HPE", "Début du traitement du reporting HP");
                }
                else
                {
                    SqlConnection con5 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                    con5.Open();
                    string myquery = "CREATE TABLE [" + textBox5.Text + "HDR_I] (FileLayoutID Varchar(50) ,FileLayoutVersion   Varchar (50) ,PartnerId   Varchar (50) ,PartnerName   Varchar (50) ,Partnercontact   Varchar (50) ,PartnerReferenceDepotID   Varchar (50) ,FileCreationDate   Varchar (50) ,FileCreationTime  Varchar (50) ,PeriodStartDate   Varchar (50) ,PeriodEndDate   Varchar (50) ,FileSequenceNumber   Varchar (50) PRIMARY KEY ,TestFileIndicator   Varchar (50))";
                    SqlCommand Cmd_HDR = new SqlCommand(myquery, con5);
                    Cmd_HDR.ExecuteNonQuery();
                    con5.Close();
                    /* MessageBox.Show("OK HDR_I"); */
                    richTextBox3.Text = "<Création> : " + "OK HDR_I";
                    SendEmail(" notifhp@disway.com", "notifhp@disway.com", "Alerte HPI", "Début du traitement du reporting HP");
                }
                return true;
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
                SqlConnection con4 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                con4.Open();
                string myquery2 = " insert into [Incidents_HP] ([Numero Etat HP] ,[Incident]) values (" + textBox5.Text.ToString() + ",'" + e.Message.ToString().Replace("'", "") + "')";
                SqlCommand Incident = new SqlCommand(myquery2, con4);
                Incident.ExecuteNonQuery();
                con4.Close();
                if (codeHP == CodeHPE)
                {
                    /*Notification_Mail*/
                    SendEmail(" notifhp@disway.com", "notifhp@disway.com", "Alerte HPE", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de HDR_E: " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                else 
                {
                    /*Notification_Mail*/
                    SendEmail(" notifhp@disway.com", "notifhp@disway.com", "Alerte HPI", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de HDR_I: " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                return false;
            }
        }
        /*Fonction de remplissage de la table HDR_E */
        private bool Remplissage_HDR(string codeHP)
        {

            /* Selection des informations sur le fichier à créer de la table Informations dans la BD Reporting-Hp-Config */
            SqlConnection Con_rmp_HDR = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
            Con_rmp_HDR.Open();
            SqlCommand cmd_FILELAYOUTID = new SqlCommand("select Valeur_information from Informations where Type_information='FILELAYOUTID'", Con_rmp_HDR);
            string FileLayoutId = (string)cmd_FILELAYOUTID.ExecuteScalar();
            SqlCommand cmd_FileLayoutVersion = new SqlCommand("select Valeur_information from Informations where Type_information='FILELAYOUTVERSION'", Con_rmp_HDR);
            string FileLayoutVersion = (string)cmd_FileLayoutVersion.ExecuteScalar();
            SqlCommand cmd_PartnerID = new SqlCommand("select Valeur_information from Informations where Type_information='PARTNERID'", Con_rmp_HDR);
            string PartnerID = (string)cmd_PartnerID.ExecuteScalar();
            SqlCommand cmd_PartnerName = new SqlCommand("select Valeur_information from Informations where Type_information='PARTNERNAME'", Con_rmp_HDR);
            string PartnerName = (string)cmd_PartnerName.ExecuteScalar();
            SqlCommand cmd_PARTNERREFERENCEDEPOTID = new SqlCommand("select Valeur_information from Informations where Type_information='PARTNERREFERENCEDEPOTID'", Con_rmp_HDR);
            string PARTNERREFERENCEDEPOTID = (string)cmd_PARTNERREFERENCEDEPOTID.ExecuteScalar();
            SqlCommand cmd_PARTNERContact = new SqlCommand("select Valeur_information from Informations where Type_information='PARTNERCONTACT'", Con_rmp_HDR);
            string PARTNERContact = (string)cmd_PARTNERContact.ExecuteScalar();
            Con_rmp_HDR.Close();
            /* Remplissage HDR_E */
            String Time_Creation_File = DateTime.Now.ToString("HH:mm"); /* Ligne a modifier */
           /* textBox1.Text = "";
            textBox2.Text = ""; 
            /*
            String Date_Creation_File = DateTime.Now.ToShortDateString().ToString("ddMMyyyy");*/

            String Date_Creation_File = DateTime.Now.ToString("yyyyMMdd");
            dateTimePicker1.CustomFormat = "yyyyMMdd";
            dateTimePicker2.CustomFormat = "yyyyMMdd";
            SqlConnection Con_StockageFiles = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
            Con_StockageFiles.Open();
            /* Rest a remplir Now par l'heure de création du fichier */
            if (comboBox1.SelectedIndex == 0)
            {
                if (codeHP == CodeHPE)
                {
                    string Myquery_HDR = "insert into  [" + textBox5.Text + "HDR_E] ( [FileLayoutID], [FileLayoutVersion], [PartnerId], [PartnerName], [Partnercontact], [PartnerReferenceDepotID], [FileCreationDate], [FileCreationTime], [PeriodStartDate], [PeriodEndDate], [FileSequenceNumber], [TestFileIndicator] ) values ('" + FileLayoutId + "','" + FileLayoutVersion + "' ,'" + Insertion_Espaces_HDR_9(PartnerID) + "','" + Insertion_Espaces_HDR_35(PartnerName) + "','" + Insertion_Espaces_HDR_35(PARTNERContact) + "','" + Insertion_Espaces_HDR_35(PARTNERREFERENCEDEPOTID) + "', '" + Insertion_Espaces_HDR_8(Date_Creation_File) + "', '" + Insertion_Espaces_HDR_4(Time_Creation_File.Replace(":", "")) + "', '" + Insertion_Espaces_HDR_8(dateTimePicker1.Text.Replace("-", "")) + "','" + Insertion_Espaces_HDR_8(dateTimePicker2.Text.Replace("-", "")) + "' ,'" + Insertion_ZERO_5DIG(textBox5.Text) + "','0')";
                    SqlCommand Cmd_Remp_HDR = new SqlCommand(Myquery_HDR, Con_StockageFiles);
                    Cmd_Remp_HDR.ExecuteNonQuery();
                    Con_StockageFiles.Close();
                    dateTimePicker1.CustomFormat = "  dd-MM-yyyy";
                    dateTimePicker2.CustomFormat = "  dd-MM-yyyy";
                    richTextBox4.Text = "<Remplissage> : " + "OK HDR_E";
                }
                else 
                {
                    string Myquery_HDR = "insert into  [" + textBox5.Text + "HDR_I] ( [FileLayoutID], [FileLayoutVersion], [PartnerId], [PartnerName], [Partnercontact], [PartnerReferenceDepotID], [FileCreationDate], [FileCreationTime], [PeriodStartDate], [PeriodEndDate], [FileSequenceNumber], [TestFileIndicator] ) values ('" + FileLayoutId + "','" + FileLayoutVersion + "' ,'" + Insertion_Espaces_HDR_9(PartnerID) + "','" + Insertion_Espaces_HDR_35(PartnerName) + "','" + Insertion_Espaces_HDR_35(PARTNERContact) + "','" + Insertion_Espaces_HDR_35(PARTNERREFERENCEDEPOTID) + "', '" + Insertion_Espaces_HDR_8(Date_Creation_File) + "', '" + Insertion_Espaces_HDR_4(Time_Creation_File.Replace(":", "")) + "', '" + Insertion_Espaces_HDR_8(dateTimePicker1.Text.Replace("-", "")) + "','" + Insertion_Espaces_HDR_8(dateTimePicker2.Text.Replace("-", "")) + "' ,'" + Insertion_ZERO_5DIG(textBox5.Text) + "','0')";
                    SqlCommand Cmd_Remp_HDR = new SqlCommand(Myquery_HDR, Con_StockageFiles);
                    Cmd_Remp_HDR.ExecuteNonQuery();
                    Con_StockageFiles.Close();
                    dateTimePicker1.CustomFormat = "  dd-MM-yyyy";
                    dateTimePicker2.CustomFormat = "  dd-MM-yyyy";
                    richTextBox4.Text = "<Remplissage> : " + "OK HDR_I";
                } ///End IF : vérification du Code HP
                return true;
            }
            else
            {
                if (codeHP == CodeHPE)
                {
                    string Myquery_HDR = "insert into  [" + textBox5.Text + "HDR_E] ( [FileLayoutID], [FileLayoutVersion], [PartnerId], [PartnerName], [Partnercontact], [PartnerReferenceDepotID], [FileCreationDate], [FileCreationTime], [PeriodStartDate], [PeriodEndDate], [FileSequenceNumber], [TestFileIndicator] ) values ('" + FileLayoutId + "','" + FileLayoutVersion + "' ,'" + Insertion_Espaces_HDR_9(PartnerID) + "','" + Insertion_Espaces_HDR_35(PartnerName) + "','" + Insertion_Espaces_HDR_35(PARTNERContact) + "','" + Insertion_Espaces_HDR_35(PARTNERREFERENCEDEPOTID) + "', '" + Insertion_Espaces_HDR_8(Date_Creation_File) + "', '" + Insertion_Espaces_HDR_4(Time_Creation_File.Replace(":", "")) + "', '" + Insertion_Espaces_HDR_8(dateTimePicker1.Text.Replace("-", "")) + "','" + Insertion_Espaces_HDR_8(dateTimePicker2.Text.Replace("-", "")) + "' ,'" + Insertion_ZERO_5DIG(textBox5.Text) + "','1')";
                    SqlCommand Cmd_Remp_HDR = new SqlCommand(Myquery_HDR, Con_StockageFiles);
                    Cmd_Remp_HDR.ExecuteNonQuery();
                    Con_StockageFiles.Close();
                    richTextBox4.Text = "<Remplissage> : " + "OK HDR_E";
                }
                else
                {
                    string Myquery_HDR = "insert into  [" + textBox5.Text + "HDR_I] ( [FileLayoutID], [FileLayoutVersion], [PartnerId], [PartnerName], [Partnercontact], [PartnerReferenceDepotID], [FileCreationDate], [FileCreationTime], [PeriodStartDate], [PeriodEndDate], [FileSequenceNumber], [TestFileIndicator] ) values ('" + FileLayoutId + "','" + FileLayoutVersion + "' ,'" + Insertion_Espaces_HDR_9(PartnerID) + "','" + Insertion_Espaces_HDR_35(PartnerName) + "','" + Insertion_Espaces_HDR_35(PARTNERContact) + "','" + Insertion_Espaces_HDR_35(PARTNERREFERENCEDEPOTID) + "', '" + Insertion_Espaces_HDR_8(Date_Creation_File) + "', '" + Insertion_Espaces_HDR_4(Time_Creation_File.Replace(":", "")) + "', '" + Insertion_Espaces_HDR_8(dateTimePicker1.Text.Replace("-", "")) + "','" + Insertion_Espaces_HDR_8(dateTimePicker2.Text.Replace("-", "")) + "' ,'" + Insertion_ZERO_5DIG(textBox5.Text) + "','1')";
                    SqlCommand Cmd_Remp_HDR = new SqlCommand(Myquery_HDR, Con_StockageFiles);
                    Cmd_Remp_HDR.ExecuteNonQuery();
                    Con_StockageFiles.Close();
                    richTextBox4.Text = "<Remplissage> : " + "OK HDR_I";
                }///End IF : vérification du Code HP
                return true;
            }

        }
        /****** ZHA 15/06/2015 
         * Fonction de création de la table SNT */
        private bool Creation_SNT_Table(string codeHP)
        {
            try
            {
                if (codeHP == "HPE")//CodeHPE)
                {
                    SqlConnection con5 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                    con5.Open();
                    string myquery = "CREATE TABLE [" + textBox5.Text + "SNT_E] ([Bundle ID] Varchar(10),[Bundle quantity] Varchar(10),[Product number]  Varchar(50),[Product quantity]  Varchar(10),[HP serial number]   Varchar(8000),[Deal ID]   Varchar(30),[End customer name]  Varchar(10),[Claim reference]  Varchar(30),[Invoice or packing number]   Varchar(30),[Shipment or sell-out date]  Varchar(30),[2nd tier reseller name]  Varchar(50),[2nd tier reseller ID] Varchar(10),[2nd tier reseller order reference]  Varchar(10),[List price]  Varchar(30),[Net price]  Varchar(30),[Net purchase price]  Varchar(30),[Deal net price]  Varchar(30),[Deal discount %]  Varchar(30))";
                    SqlCommand Cmd_HDR = new SqlCommand(myquery, con5);
                    Cmd_HDR.ExecuteNonQuery();
                    con5.Close();
                    /* MessageBox.Show("OK SNT_E"); */
                    richTextBox3.Text = "<Création> : " + "OK SNT_E";
                }
                else
                {
                    SqlConnection con5 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                    con5.Open();
                    string myquery = "CREATE TABLE [" + textBox5.Text + "SNT_I] ([Bundle ID] Varchar(10),[Bundle quantity] Varchar(10),[Product number]  Varchar(50),[Product quantity]  Varchar(10),[HP serial number]   Varchar(8000),[Deal ID]   Varchar(30),[End customer name]  Varchar(10),[Claim reference]  Varchar(30),[Invoice or packing number]   Varchar(30),[Shipment or sell-out date]  Varchar(30),[2nd tier reseller name]  Varchar(50),[2nd tier reseller ID] Varchar(10),[2nd tier reseller order reference]  Varchar(10),[List price]  Varchar(30),[Net price]  Varchar(30),[Net purchase price]  Varchar(30),[Deal net price]  Varchar(30),[Deal discount %]  Varchar(30))";
                    SqlCommand Cmd_HDR = new SqlCommand(myquery, con5);
                    Cmd_HDR.ExecuteNonQuery();
                    con5.Close();
                    /* MessageBox.Show("OK SNT_I"); */
                    richTextBox3.Text = "<Création> : " + "OK SNT_I";
                }
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                SqlConnection con4 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                con4.Open();
                string myquery2 = " insert into [Incidents_HP] ([Numero Etat HP] ,[Incident]) values (" + textBox5.Text.ToString() + ",'" + e.Message.ToString().Replace("'", "") + "')";
                SqlCommand Incident = new SqlCommand(myquery2, con4);
                Incident.ExecuteNonQuery();
                con4.Close();
                if (codeHP == CodeHPE)
                {
                    /*Notification_Mail*/
                    SendEmail(" notifhp@disway.com", "notifhp@disway.com", "Alerte HPE", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de SNT_E: " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                else
                {
                    /*Notification_Mail*/
                    SendEmail(" notifhp@disway.com", "notifhp@disway.com", "Alerte HPI", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de SNT_I: " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                return false;
            }
        }
        /*Fonction de remplissage de la table SNT */
        private bool Remplissage_SNT(string codeHP,string BL,string LBL)
        {
            if (codeHP == CodeHPE)
            {
                /* Selection des codes familles HPE *
                SqlConnection Con_FHP = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
                Con_FHP.Open();
                String Query = "select DISTINCT [code_Famille] from Famille_HP_New where Cartes='HPE'";
                SqlCommand cmd_FamilleHP = new SqlCommand(Query, Con_FHP);
                SqlDataReader Reader = cmd_FamilleHP.ExecuteReader();
                string FHP = "";
                string[] THP = new string[100];
                int Count = 0;
                while (Reader.Read())
                {
                    THP[Count] = Reader.GetString(0);
                    Count++;
                }
                Reader.Close();
                for (int j = 0; j < Count; j++)
                {
                    if (j < Count - 1)
                        FHP += "'" + THP[j] + "',";
                    else
                        FHP += "'" + THP[j] + "'";
                }
                Con_FHP.Close();*/
                //int Count = 0;
                /* Selection des informations sur le fichier à créer de la table Informations dans la BD Disway */
                SqlConnection Con_rmp_SNT = new SqlConnection("Data Source=WAYBI;Initial Catalog=Disway;Integrated Security=True");
                Con_rmp_SNT.Open();
                SqlCommand cmd_SNT = new SqlCommand("SELECT [Bundle ID],[Bundle quantity], [Product number], [Product quantity],[HP serial number], [Deal ID], [End customer name],[Claim reference], [Invoice or packing number],[Shipment or sell-out date],[2nd tier reseller name], [2nd tier reseller ID], [2nd tier reseller order reference], [List price], [Net price], [Net purchase price], [Deal net price],[Deal discount %] FROM ____Claim_HP WHERE ([N° BL]='" + BL + "') and ([N° ligne BL]='" + LBL + "')", Con_rmp_SNT);
                cmd_SNT.CommandTimeout = 0;
                cmd_SNT.ExecuteNonQuery();
                SqlDataReader Reader_SNT = cmd_SNT.ExecuteReader();
                DataTable SNT_DT = new DataTable();
                SNT_DT.Load(Reader_SNT);
                int Count_myquery_SNT = SNT_DT.Rows.Count;
                Reader_SNT = cmd_SNT.ExecuteReader();
                /*Count = 0;
                Count = Count_myquery_SNT + 1;
                string[] Z = new string[Count];
                string[] Y = new string[Count];
                //int b = 0;*/
                /* On récupère chaque ligne dans des tableaux */
                SqlConnection Con_StockageFiles = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                Con_StockageFiles.Open();
                while (Reader_SNT.Read())
                {

                    string Bundle_ID = Reader_SNT.GetString(0);
                    string Bundle_quantity = Reader_SNT.GetString(1);
                    string Product_number = Reader_SNT.GetString(2);
                    string Product_quantity = Reader_SNT.GetString(3);
                     string HP_serial_number = null;
                    if (!Reader_SNT.IsDBNull(4)) HP_serial_number = Reader_SNT.GetString(4);
                    string Deal_ID = Reader_SNT.GetString(5);
                    string End_customer_name = Reader_SNT.GetString(6);
                    string Claim_reference = Reader_SNT.GetString(7);
                    string Invoice_or_packing_number = Reader_SNT.GetString(8);
                    string Shipment_or_sell_out_date = Reader_SNT.GetDateTime(9).ToString();//Reader_SNT.GetString(9); 
                    string tier_reseller_name = Reader_SNT.GetString(10);
                    string tier_reseller_ID = Reader_SNT.GetString(11);
                    string tier_reseller_order_reference = Reader_SNT.GetString(12);
                    string List_price = Reader_SNT.GetString(13);
                    string Net_price = Reader_SNT.GetString(14);
                    string Net_purchase_price = Reader_SNT.GetString(15);
                    string Deal_net_price = Reader_SNT.GetString(16);
                    string Deal_discount_pour = Reader_SNT.GetString(17);

                    string Myquery_SNT = "insert into  [" + textBox5.Text + "SNT_E] ([Product number],[Product quantity],[HP serial number],[Deal ID],[Claim reference],[Invoice or packing number],[Shipment or sell-out date],[2nd tier reseller name], [2nd tier reseller ID],[List price],[Net price],[Net purchase price],[Deal net price],[Deal discount %]) values ('" + Product_number + "','" + Product_quantity + "','" + HP_serial_number + "','" + Deal_ID + "','" + Claim_reference + "','" + Invoice_or_packing_number + "','" + Shipment_or_sell_out_date + "','" + tier_reseller_name + "','" + tier_reseller_ID + "','" + List_price + "','" + Net_price + "','" + Net_purchase_price + "','" + Deal_net_price + "','" + Deal_discount_pour + "')";
                    SqlCommand Cmd_Remp_SNT = new SqlCommand(Myquery_SNT, Con_StockageFiles);
                    Cmd_Remp_SNT.ExecuteNonQuery();
                }
                Reader_SNT.Close();
                Con_rmp_SNT.Close();
                Con_StockageFiles.Close();
                //dateTimePicker1.CustomFormat = "  dd-MM-yyyy";
                //dateTimePicker2.CustomFormat = "  dd-MM-yyyy";
                richTextBox4.Text = "<Remplissage> : " + "OK SNT_E";
            }
            else
            {
                /* Selection des codes familles HPI *
                SqlConnection Con_FHP = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
                Con_FHP.Open();
                String Query = "select DISTINCT [code_Famille] from Famille_HP_New where Cartes='HPI'";
                SqlCommand cmd_FamilleHP = new SqlCommand(Query, Con_FHP);
                SqlDataReader Reader = cmd_FamilleHP.ExecuteReader();
                string FHP = "";
                string[] THP = new string[100];
                int Count = 0;
                while (Reader.Read())
                {
                    THP[Count] = Reader.GetString(0);
                    Count++;
                }
                Reader.Close();
                for (int j = 0; j < Count; j++)
                {
                    if (j < Count - 1)
                        FHP += "'" + THP[j] + "',";
                    else
                        FHP += "'" + THP[j] + "'";
                }
                Con_FHP.Close();*/
                //int Count = 0;
                /* Selection des informations sur le fichier à créer de la table Informations dans la BD Disway */
                SqlConnection Con_rmp_SNT = new SqlConnection("Data Source=WAYBI;Initial Catalog=Disway;Integrated Security=True");
                Con_rmp_SNT.Open();
                SqlCommand cmd_SNT = new SqlCommand("SELECT [Bundle ID],[Bundle quantity], [Product number], [Product quantity],[HP serial number], [Deal ID], [End customer name],[Claim reference], [Invoice or packing number],[Shipment or sell-out date],[2nd tier reseller name], [2nd tier reseller ID], [2nd tier reseller order reference], [List price], [Net price], [Net purchase price], [Deal net price],[Deal discount %] FROM ____Claim_HP WHERE ([N° BL]='" + BL + "') and ([N° ligne BL]='" + LBL + "')", Con_rmp_SNT);
                cmd_SNT.CommandTimeout = 0;
                cmd_SNT.ExecuteNonQuery();
                SqlDataReader Reader_SNT = cmd_SNT.ExecuteReader();
                DataTable SNT_DT = new DataTable();
                SNT_DT.Load(Reader_SNT);
                int Count_myquery_SNT = SNT_DT.Rows.Count;
                Reader_SNT = cmd_SNT.ExecuteReader();
                /*Count = 0;
                Count = Count_myquery_SNT + 1;
                //string[] Z = new string[Count];
                //string[] Y = new string[Count];
                //int b = 0;*/
                /* On récupère chaque ligne dans des tableaux */
                SqlConnection Con_StockageFiles = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                Con_StockageFiles.Open();
                while (Reader_SNT.Read())
                {
                    string Bundle_ID = Reader_SNT.GetString(0);
                    string Bundle_quantity = Reader_SNT.GetString(1);
                    string Product_number = Reader_SNT.GetString(2);
                    string Product_quantity = Reader_SNT.GetString(3);
                    string HP_serial_number = null;
                    if (!Reader_SNT.IsDBNull(4)) HP_serial_number = Reader_SNT.GetString(4);
                    string Deal_ID = Reader_SNT.GetString(5);
                    string End_customer_name = Reader_SNT.GetString(6);
                    string Claim_reference = Reader_SNT.GetString(7);
                    string Invoice_or_packing_number = Reader_SNT.GetString(8);
                    string Shipment_or_sell_out_date = Reader_SNT.GetDateTime(9).ToString();
                    string tier_reseller_name = Reader_SNT.GetString(10);
                    string tier_reseller_ID = Reader_SNT.GetString(11);
                    string tier_reseller_order_reference = Reader_SNT.GetString(12);
                    string List_price = Reader_SNT.GetString(13);
                    string Net_price = Reader_SNT.GetString(14);
                    string Net_purchase_price = Reader_SNT.GetString(15);
                    string Deal_net_price = Reader_SNT.GetString(16);
                    string Deal_discount_pour = Reader_SNT.GetString(17);

                    string Myquery_SNT = "insert into  [" + textBox5.Text + "SNT_I] ([Product number],[Product quantity],[HP serial number],[Deal ID],[Claim reference],[Invoice or packing number],[Shipment or sell-out date],[2nd tier reseller name],[2nd tier reseller ID],[List price],[Net price],[Net purchase price],[Deal net price],[Deal discount %]) values ('" + Product_number + "','" + Product_quantity + "','" + HP_serial_number + "','" + Deal_ID + "','" + Claim_reference + "','" + Invoice_or_packing_number + "','" + Shipment_or_sell_out_date + "','" + tier_reseller_name + "','" + tier_reseller_ID + "','" + List_price + "','" + Net_price + "','" + Net_purchase_price + "','" + Deal_net_price + "','" + Deal_discount_pour + "')";
                    SqlCommand Cmd_Remp_SNT = new SqlCommand(Myquery_SNT, Con_StockageFiles);
                    Cmd_Remp_SNT.ExecuteNonQuery();
                }
                Reader_SNT.Close();
                Con_rmp_SNT.Close();
                Con_StockageFiles.Close();
                dateTimePicker1.CustomFormat = "  dd-MM-yyyy";
                dateTimePicker2.CustomFormat = "  dd-MM-yyyy";
                richTextBox4.Text = "<Remplissage> : " + "OK SNT_I";
            }///End IF : vérification du Code HP
            return true;
        }

        /* Fonction de remplissage de la fonction SIT */
        private bool Remplissage_SIT(string codeHP)
        {
            try
            {
                if (codeHP == CodeHPE)
                {
                    SqlConnection con_SITE = new SqlConnection("Data Source=WAYBI;Initial Catalog=Disway;Integrated Security=True");
                    con_SITE.Open();
                    SqlCommand Cmd = new SqlCommand("Remplir_SITE_NEW", con_SITE);
                    Cmd.CommandType = CommandType.StoredProcedure;
                    Cmd.Parameters.AddWithValue("@NSeq", textBox5.Text);
                    Cmd.Parameters.AddWithValue("@Date_debut", dateTimePicker1.Text);
                    Cmd.Parameters.AddWithValue("@Date_fin", dateTimePicker2.Text);
                    Cmd.Parameters.AddWithValue("@CodeHP", codeHP);
                    SqlParameter sqlParam = new SqlParameter("@Result", DbType.Boolean);
                    sqlParam.Direction = ParameterDirection.Output;
                    Cmd.Parameters.Add(sqlParam);
                    Cmd.CommandTimeout = 10000;
                    //Cmd.ExecuteNonQuery();
                    SqlDataReader rdr = Cmd.ExecuteReader();
                    con_SITE.Close();
                    richTextBox4.Text = richTextBox4.Text + "  / " + "OK SIT_E";
                    Application.DoEvents(); 
                    return Convert.ToBoolean(Cmd.Parameters["@Result"].Value);
                }
                else
                {
                    SqlConnection con_SITI = new SqlConnection("Data Source=WAYBI;Initial Catalog=Disway;Integrated Security=True");
                    con_SITI.Open();
                    SqlCommand Cmd = new SqlCommand("Remplir_SITI_AUR", con_SITI);
                    Cmd.CommandType = CommandType.StoredProcedure;
                    Cmd.Parameters.AddWithValue("@NSeq", textBox5.Text);
                    Cmd.Parameters.AddWithValue("@Date_debut", dateTimePicker1.Text);
                    Cmd.Parameters.AddWithValue("@Date_fin", dateTimePicker2.Text);
                    Cmd.Parameters.AddWithValue("@CodeHP", codeHP);
                    SqlParameter sqlParam = new SqlParameter("@Result", DbType.Boolean);
                    sqlParam.Direction = ParameterDirection.Output;
                    Cmd.Parameters.Add(sqlParam);
                    Cmd.CommandTimeout = 10000;
                    SqlDataReader rdr = Cmd.ExecuteReader(); 
                    //Cmd.ExecuteNonQuery();
                    con_SITI.Close();
                    richTextBox4.Text = richTextBox4.Text + "  / " + "OK SIT_I";
                    Application.DoEvents(); 
                    return Convert.ToBoolean(Cmd.Parameters["@Result"].Value);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                SqlConnection con4 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                con4.Open();
                string myquery2 = " insert into [Incidents_HP] ([Numero Etat HP] ,[Incident]) values (" + textBox5.Text.ToString() + ",'" + e.Message.ToString().Replace("'", "") + "')";
                SqlCommand Incident = new SqlCommand(myquery2, con4);
                Incident.ExecuteNonQuery();
                con4.Close();
                if (codeHP == CodeHPE)
                {
                    MessageBox.Show("Remplissage SIT_E not OK !");
                    MessageBox.Show(e.Message);
                    /*Notification_Mail*/
                    SendEmail("notifhp@disway.com", "notifhp@disway.com", "Alerte HPE", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de SIT_E: " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                else
                {
                    MessageBox.Show("Remplissage SIT_I not OK !");
                    MessageBox.Show(e.Message);
                    /*Notification_Mail*/
                    SendEmail("notifhp@disway.com", "notifhp@disway.com", "Alerte HPI", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de SIT_I: " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                return false;
            }
        }
                
        /*Fonction de traitement de la table PRS_E */
        private bool Traitement_PRS(string codeHP)
        {
            try
            {
                if (codeHP == CodeHPE)
                {
                    /* Selection des lignes de la table PRS_E */
                    SqlConnection Con_HP = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;MultipleActiveResultSets=True;Integrated Security=True");
                    Con_HP.Open();
                    string Query_HP_PRS = "select * from [Reporting-Hp-StorageFiles].[dbo].[" + textBox5.Text + "PRS_E]";
                    string Query_PRS_Count = "select count(*) from [Reporting-Hp-StorageFiles].[dbo].[" + textBox5.Text + "PRS_E]";
                    SqlCommand Cmd_Count_PRS = new SqlCommand(Query_PRS_Count, Con_HP);
                    int Count_PRS = (int)Cmd_Count_PRS.ExecuteScalar();


                    /* Selection des composants KIt HP */
                    SqlConnection Con_KIT = new SqlConnection("Data Source=WAYBI;Initial Catalog=DISWAY;MultipleActiveResultSets=True;Integrated Security=True");
                    Con_KIT.Open();
                    string Query_KIT = "select count(*) FROM  [__Composants_Kit_HP]";
                    SqlCommand Cmd_Count_KIT = new SqlCommand(Query_KIT, Con_KIT);
                    int Count_KIT = (int)Cmd_Count_KIT.ExecuteScalar();

                    int r = 0;//, m = 0;
                    bool Trouve;
                    string[] R2 = new string[2000];
                    string[] R0 = new string[2000];
                    string[] R4 = new string[2000];
                    string[] R7 = new string[2000];
                    /***
                     * ZHA 29/05/2015 : HP OPG
                     */
                    string[] R11 = new string[2000];
                    string[] R12 = new string[2000];
                    string[] R13 = new string[2000];
                    string[] R14 = new string[2000];
                     
                    string[] DS = new string[2000];
                    string[] DS1 = new string[2000];
                    string[] DS3 = new string[2000];
                    string[] DS4 = new string[2000];
                    string[] DS0 = new string[2000];
                    int[] Qte = new int[2000];

                    SqlCommand Cmd_PRS = new SqlCommand(Query_HP_PRS, Con_HP);
                    SqlDataReader Reader_PRS = Cmd_PRS.ExecuteReader();

                    string BundleID = "";
                    string ProdSerialID = "";
                    string PartPurchOrderID = "POHP";
                    string HPEInvoiceNumber = "";
                    string EndUserID = "";
                    string ShipToCustID = "";
                    string OriginCountry = "MA";
                    string DropShipFlag = "N";
                    string UpFrontOPG2 = "";
                    string BackendOPG2 = "";
                    string UpFrontOPG3 = "";
                    string BackendOPG3 = "";
                    string BackendOPG4 = "";
                    string DealRegID1 = "";
                    string DealRegID2 = "";
                    string PartInternTransID = "";
                    string OriginalHPETransNumber = "";
                    string PartPurshPrice = "";
                    string ExtendNetCostAfterRebate = "";
                    string TerritoryManager = "";

                    while (Reader_PRS.Read())
                    {
                        Trouve = false;
                        R2[r] = Reader_PRS[2].ToString(); R4[r] = Reader_PRS[4].ToString(); R0[r] = Reader_PRS[0].ToString();
                        R7[r] = Reader_PRS[7].ToString();
                        /***
                         * ZHA 29/05/2015 : HP OPG*/
                        R11[r] = Reader_PRS[11].ToString();
                        R12[r] = Reader_PRS[12].ToString();
                        R13[r] = Reader_PRS[13].ToString();
                        R14[r] = Reader_PRS[14].ToString();
                        
                        string Query_HP_Comp_KIT = "SELECT  [N° kit],[N° composant],[Description],CONVERT(int ,ISNULL([Quantité] ,0) ),[Ref] ,[Manufacturer Code]FROM [__Composants_Kit_HP]";
                        SqlCommand Cmd_Comp_KIT = new SqlCommand(Query_HP_Comp_KIT, Con_KIT);
                        SqlDataReader Reader_Comp_KIT = Cmd_Comp_KIT.ExecuteReader();
                        while (Reader_Comp_KIT.Read())
                        {
                            DS0[r] = Insertion_Espaces_SIT_20(Reader_Comp_KIT[0].ToString());
                            /* DS1[r] = Insertion_Espaces_SIT_20(Reader_Comp_KIT[1].ToString());*/
                            DS1[r] = Reader_Comp_KIT[1].ToString();
                            DS3[r] = Reader_Comp_KIT[3].ToString();
                            DS4[r] = Insertion_Espaces_SIT_20(Reader_Comp_KIT[4].ToString());
     
                            if (R2[r] == DS4[r])
                            {
                                String RS_TMP = "select * from [Reporting-Hp-StorageFiles].[dbo].[" + textBox5.Text + "PRS_E] where [HPProductNumber] = '" + DS1[r] + "' and [SoldToCustomerID] ='" + R0[r] + "'";
                                SqlCommand Cmd_RS = new SqlCommand(RS_TMP, Con_HP);
                                SqlDataReader Reader_RS = Cmd_RS.ExecuteReader();

                                /* MessageBox.Show("Commun trouvé '"+ R2[r]+"'"); */

                                string[] TMP0 = new string[1200];
                                string[] TMP2 = new string[1200];
                                string[] TMP3 = new string[1200];
                                string[] TMP4 = new string[1200];
                                string[] TMP5 = new string[1200];
                                string[] TMP8 = new string[1200];
                                string[] TMP9 = new string[1200];
                                while (Reader_RS.Read())
                                {
                                    TMP0[r] = Reader_RS[0].ToString();
                                    TMP2[r] = Reader_RS[2].ToString();
                                    TMP3[r] = Reader_RS[3].ToString();
                                    TMP4[r] = Reader_RS[4].ToString();
                                    TMP5[r] = Reader_RS[5].ToString();
                                    TMP8[r] = Reader_RS[8].ToString();
                                    TMP9[r] = Reader_RS[9].ToString();
                                }

                                if (Reader_RS.HasRows)
                                {


                                    Qte[r] = (int.Parse(R4[r]) * int.Parse(DS3[r])) + int.Parse(TMP4[r]);
                                    string Query_MAJ_PRS = "update  [" + textBox5.Text + "PRS_E] set [DetailedSellout] ='" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(Qte[r].ToString()))) + "' where [HPProductNumber] = '" + TMP2[r] + "'and [SoldToCustomerID] ='" + TMP0[r] + "'";
                                    SqlCommand MAJ_PRS = new SqlCommand(Query_MAJ_PRS, Con_HP);
                                    MAJ_PRS.ExecuteNonQuery();
                                }
                                else
                                {
                                    Qte[r] = (int.Parse(R4[r]) * int.Parse(DS3[r]));
                                    //String Query_INS_PRS = " insert into  [" + textBox5.Text + "PRS_E]  ([SoldToCustomerID] ,[InvoicetoCustomerID],[HPProductNumber] ,[PartnerProductNumber],[DetailedSellout],[SpecialPrincingreference],[InvoiceNumber],[InvoiceDate],[ShipmentDate],[InvoiceNetAmount],[CurrencyCode]) values ('" + Insertion_Espaces_PRS_17(R0[r]) + "', '" + Insertion_Espaces_PRS_17(R0[r]) + "', '" + Insertion_Espaces_PRS_20(DS1[r]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(Qte[r].ToString()))) + "','" + Insertion_Espaces_PRS_6("") + "' ,'" + Insertion_Espaces_PRS_30("") + "', '" + Insertion_Espaces_PRS_8(R7[r].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8("") + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "')";
                                    /***
                                     * ZHA 29/05/2015 : HP OPG
                                     **/
                                    
                                    //String Query_INS_PRS = " insert into  [" + textBox5.Text + "PRS_E]  ([SoldToCustomerID] ,[InvoicetoCustomerID],[HPProductNumber] ,[PartnerProductNumber],[DetailedSellout],[SpecialPrincingreference],[InvoiceNumber],[InvoiceDate],[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2]) values ('" + Insertion_Espaces_PRS_17(R0[r]) + "', '" + Insertion_Espaces_PRS_17(R0[r]) + "', '" + Insertion_Espaces_PRS_20(DS1[r]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(Qte[r].ToString()))) + "','" + Insertion_Espaces_PRS_6("") + "' ,'" + Insertion_Espaces_PRS_30("") + "', '" + Insertion_Espaces_PRS_8(R7[r].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8("") + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "', '" + R11[r] + "', '" + R12[r] + "', '" + R13[r] + "', '" + R14[r] +"')";
                                    String Query_INS_PRS = " insert into  [" + textBox5.Text + "PRS_E]  ([SoldToCustomerID] ,[InvoicetoCustomerID],[HPProductNumber] ,[PartnerProductNumber],[DetailedSellout],[SpecialPrincingreference],[InvoiceNumber],[InvoiceDate],[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2],[BundleID], [ProdSerialID], [PartPurchOrderID], [HPEInvoiceNumber], [EndUserID], [ShipToCustID], [OriginCountry], [DropShipFlag], [UpFrontOPG2], [BackendOPG2], [UpFrontOPG3], [BackendOPG3], [BackendOPG4], [DealRegID1], [DealRegID2], [PartInternTransID], [OriginalHPETransNumber], [PartPurshPrice], [ExtendNetCostAfterRebate], [TerritoryManager]) values ('" + Insertion_Espaces_PRS_17(R0[r]) + "', '" + Insertion_Espaces_PRS_17(R0[r]) + "', '" + Insertion_Espaces_PRS_20(DS1[r]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(Qte[r].ToString()))) + "','" + Insertion_Espaces_PRS_6("") + "' ,'" + Insertion_Espaces_PRS_30("") + "', '" + Insertion_Espaces_PRS_8(R7[r].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8("") + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "', '" + R11[r] + "', '" + R12[r] + "', '" + R13[r] + "', '" + R14[r] + "','" + BundleID.PadRight(15, ' ') + "', '" + ProdSerialID.PadRight(15, ' ') + "', '" + PartPurchOrderID.PadRight(30, ' ') + "','" + HPEInvoiceNumber.PadRight(8, ' ') + "', '" + EndUserID.PadRight(17, ' ') + "', '" + ShipToCustID.PadRight(17, ' ') + "','" + OriginCountry + "', '" + DropShipFlag + "', '" + UpFrontOPG2.PadRight(10, ' ') + "','" + BackendOPG2.PadRight(10, ' ') + "', '" + UpFrontOPG3.PadRight(10, ' ') + "',  '" + BackendOPG3.PadRight(10, ' ') + "','" + BackendOPG4.PadRight(10, ' ') + "', '" + DealRegID1.PadRight(15, ' ') + "', '" + DealRegID2.PadRight(15, ' ') + "','" + PartInternTransID.PadRight(30, ' ') + "', '" + OriginalHPETransNumber.PadRight(30, ' ') + "', '" + PartPurshPrice.PadRight(15, ' ') + "', '" + ExtendNetCostAfterRebate.PadRight(15, ' ') + "', '" + TerritoryManager.PadRight(35, ' ') + "')";

                                    SqlCommand Cmd_Remp_PRS = new SqlCommand(Query_INS_PRS, Con_HP);
                                    Cmd_Remp_PRS.ExecuteNonQuery();

                                }

                                Trouve = true;
                                Reader_RS.Close();

                            }

                        }
                        Reader_Comp_KIT.Close();
                        if (Trouve == true)
                        {
                            String Query_Delete = "delete from [" + textBox5.Text + "PRS_E]  where [HPProductNumber] = '" + R2[r] + "'and [SoldToCustomerID] = '" + R0[r] + "'";
                            SqlCommand Cmd_Del = new SqlCommand(Query_Delete, Con_HP);
                            Cmd_Del.ExecuteNonQuery();

                        }
                        r++;
                    }
                    Con_KIT.Close();
                    Con_HP.Close();
                }
                else
                {
                    /* Selection des lignes de la table PRS_I */
                    SqlConnection Con_HP = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;MultipleActiveResultSets=True;Integrated Security=True");
                    Con_HP.Open();
                    string Query_HP_PRS = "select * from [Reporting-Hp-StorageFiles].[dbo].[" + textBox5.Text + "PRS_I]";
                    string Query_PRS_Count = "select count(*) from [Reporting-Hp-StorageFiles].[dbo].[" + textBox5.Text + "PRS_I]";
                    SqlCommand Cmd_Count_PRS = new SqlCommand(Query_PRS_Count, Con_HP);
                    int Count_PRS = (int)Cmd_Count_PRS.ExecuteScalar();


                    /* Selection des composants KIt HP */
                    SqlConnection Con_KIT = new SqlConnection("Data Source=WAYBI;Initial Catalog=DISWAY;MultipleActiveResultSets=True;Integrated Security=True");
                    Con_KIT.Open();
                    string Query_KIT = "select count(*) FROM  [__Composants_Kit_HP]";
                    SqlCommand Cmd_Count_KIT = new SqlCommand(Query_KIT, Con_KIT);
                    int Count_KIT = (int)Cmd_Count_KIT.ExecuteScalar();

                    int r = 0;//, m = 0;
                    bool Trouve;
                    string[] R2 = new string[2000];
                    string[] R0 = new string[2000];
                    string[] R4 = new string[2000];
                    string[] R7 = new string[2000];
                    /***
                     * ZHA 29/05/2015 : HP OPG
                     * */
                    string[] R11 = new string[2000];
                    string[] R12 = new string[2000];
                    string[] R13 = new string[2000];
                    string[] R14 = new string[2000];

                    string[] DS = new string[2000];
                    string[] DS1 = new string[2000];
                    string[] DS3 = new string[2000];
                    string[] DS4 = new string[2000];
                    string[] DS0 = new string[2000];
                    int[] Qte = new int[2000];

                    SqlCommand Cmd_PRS = new SqlCommand(Query_HP_PRS, Con_HP);
                    SqlDataReader Reader_PRS = Cmd_PRS.ExecuteReader();

                    string BundleID1 = " ";
                    string BundleID2 = " ";
                    string OPGID3 = " ";
                    string BundleID3 = " ";
                    string OPGID4 = " ";
                    string BundleID4 = " ";
                    string OPGID5 = " ";
                    string BundleID5 = " ";
                    string OPGID6 = " ";
                    string BundleID6 = " ";
                    //string ShipToLocID = " ";
                    string HPInvoiceNo = " ";
                    string EndUserID = " ";
                    //string PartPurchPrice = " ";
                    string PartPurchPriceCC = " ";
                    //string PartRequestedRebateAmount = " ";
                    string PartnerComment = " ";
                    string PartnerReportedCBN = " ";
                    string PartnerReference = " ";
                    string PartnerInterTransID = " ";
                    string IsDropShip = " ";
                    string CustChannelPurchId = " ";
                    string PurchaseAgreement = " ";
                    string ReporterPurchOrderID = " ";
                    string SuppliesTrackID = " ";
                    string IntercompanyFlag = " ";
                    string ProdSerialIdHP = " ";

                    while (Reader_PRS.Read())
                    {
                        Trouve = false;
                        R2[r] = Reader_PRS[2].ToString(); R4[r] = Reader_PRS[4].ToString(); R0[r] = Reader_PRS[0].ToString();
                        R7[r] = Reader_PRS[7].ToString();
                        /***
                         * ZHA 29/05/2015 : HP OPG
                         * */
                        R11[r] = Reader_PRS[11].ToString();
                        R12[r] = Reader_PRS[12].ToString();
                        R13[r] = Reader_PRS[13].ToString();
                        R14[r] = Reader_PRS[14].ToString();
                        

                        string Query_HP_Comp_KIT = "SELECT  [N° kit],[N° composant],[Description],CONVERT(int ,ISNULL([Quantité] ,0) ),[Ref] ,[Manufacturer Code]FROM [__Composants_Kit_HP]";
                        SqlCommand Cmd_Comp_KIT = new SqlCommand(Query_HP_Comp_KIT, Con_KIT);
                        SqlDataReader Reader_Comp_KIT = Cmd_Comp_KIT.ExecuteReader();
                        while (Reader_Comp_KIT.Read())
                        {
                            DS0[r] = Insertion_Espaces_SIT_20(Reader_Comp_KIT[0].ToString());
                            /* DS1[r] = Insertion_Espaces_SIT_20(Reader_Comp_KIT[1].ToString());*/
                            DS1[r] = Reader_Comp_KIT[1].ToString();
                            DS3[r] = Reader_Comp_KIT[3].ToString();
                            DS4[r] = Insertion_Espaces_SIT_20(Reader_Comp_KIT[4].ToString());
                            if (R2[r] == DS4[r])
                            {
                                String RS_TMP = "select * from [Reporting-Hp-StorageFiles].[dbo].[" + textBox5.Text + "PRS_I] where [HPProductNumber] = '" + DS1[r] + "' and [SoldToCustomerID] ='" + R0[r] + "'";
                                SqlCommand Cmd_RS = new SqlCommand(RS_TMP, Con_HP);
                                SqlDataReader Reader_RS = Cmd_RS.ExecuteReader();

                                /* MessageBox.Show("Commun trouvé '"+ R2[r]+"'"); */

                                string[] TMP0 = new string[1200];
                                string[] TMP2 = new string[1200];
                                string[] TMP3 = new string[1200];
                                string[] TMP4 = new string[1200];
                                string[] TMP5 = new string[1200];
                                string[] TMP8 = new string[1200];
                                string[] TMP9 = new string[1200];

                                while (Reader_RS.Read())
                                {
                                    TMP0[r] = Reader_RS[0].ToString();
                                    TMP2[r] = Reader_RS[2].ToString();
                                    TMP3[r] = Reader_RS[3].ToString();
                                    TMP4[r] = Reader_RS[4].ToString();
                                    TMP5[r] = Reader_RS[5].ToString();
                                    TMP8[r] = Reader_RS[8].ToString();
                                    TMP9[r] = Reader_RS[9].ToString();
                                }

                                if (Reader_RS.HasRows)
                                {


                                    Qte[r] = (int.Parse(R4[r]) * int.Parse(DS3[r])) + int.Parse(TMP4[r]);
                                    string Query_MAJ_PRS = "update  [" + textBox5.Text + "PRS_I] set [DetailedSellout] ='" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(Qte[r].ToString()))) + "' where [HPProductNumber] = '" + TMP2[r] + "'and [SoldToCustomerID] ='" + TMP0[r] + "'";
                                    SqlCommand MAJ_PRS = new SqlCommand(Query_MAJ_PRS, Con_HP);
                                    MAJ_PRS.ExecuteNonQuery();
                                }
                                else
                                {
                                    Qte[r] = (int.Parse(R4[r]) * int.Parse(DS3[r]));
                                    //String Query_INS_PRS = " insert into  [" + textBox5.Text + "PRS_I]  ([SoldToCustomerID] ,[InvoicetoCustomerID],[HPProductNumber] ,[PartnerProductNumber],[DetailedSellout],[SpecialPrincingreference],[InvoiceNumber],[InvoiceDate],[ShipmentDate],[InvoiceNetAmount],[CurrencyCode]) values ('" + Insertion_Espaces_PRS_17(R0[r]) + "', '" + Insertion_Espaces_PRS_17(R0[r]) + "', '" + Insertion_Espaces_PRS_20(DS1[r]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(Qte[r].ToString()))) + "','" + Insertion_Espaces_PRS_6("") + "' ,'" + Insertion_Espaces_PRS_30("") + "', '" + Insertion_Espaces_PRS_8(R7[r].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8("") + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "')";
                                    /***
                                     * ZHA 29/05/2015 : HP OPG
                                     * */
                                    
                                    //String Query_INS_PRS = " insert into  [" + textBox5.Text + "PRS_I]  ([SoldToCustomerID] ,[InvoicetoCustomerID],[HPProductNumber] ,[PartnerProductNumber],[DetailedSellout],[SpecialPrincingreference],[InvoiceNumber],[InvoiceDate],[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2]) values ('" + Insertion_Espaces_PRS_17(R0[r]) + "', '" + Insertion_Espaces_PRS_17(R0[r]) + "', '" + Insertion_Espaces_PRS_20(DS1[r]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(Qte[r].ToString()))) + "','" + Insertion_Espaces_PRS_6("") + "' ,'" + Insertion_Espaces_PRS_30("") + "', '" + Insertion_Espaces_PRS_8(R7[r].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8("") + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "', '" + R11[r] + "', '" + R12[r] + "', '" + R13[r] + "', '" + R14[r] +"')";
                                    String Query_INS_PRS = " insert into  [" + textBox5.Text + "PRS_I]  ([SoldToCustomerID] ,[InvoicetoCustomerID],[HPProductNumber] ,[PartnerProductNumber],[DetailedSellout],[SpecialPrincingreference],[InvoiceNumber],[InvoiceDate],[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2], [BundleID1], [BundleID2], [OPGID3], [BundleID3], [OPGID4], [BundleID4], [OPGID5], [BundleID5], [OPGID6], [BundleID6], [ShipToLocID], [HPInvoiceNo], [EndUserID], [PartPurchPrice], [PartPurchPriceCC], [PartRequestedRebateAmount], [PartnerComment], [PartnerReportedCBN], [PartnerReference], [PartnerInterTransID], [IsDropShip], [CustChannelPurchId], [PurchaseAgreement], [ReporterPurchOrderID], [SuppliesTrackID], [IntercompanyFlag], [ProdSerialIdHP]) values ('" + Insertion_Espaces_PRS_17(R0[r]) + "', '" + Insertion_Espaces_PRS_17(R0[r]) + "', '" + Insertion_Espaces_PRS_20(DS1[r]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(Qte[r].ToString()))) + "','" + Insertion_Espaces_PRS_6("") + "' ,'" + Insertion_Espaces_PRS_30("") + "', '" + Insertion_Espaces_PRS_8(R7[r].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(R7[r].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "', '" + Insertion_Espaces_PRS_10(R11[r]) + "', '" + Insertion_Espaces_PRS_3(R12[r]) + "', '" + Insertion_Espaces_PRS_10(R13[r]) + "', '" + Insertion_Espaces_PRS_3(R14[r]) + "', '" + Insertion_Espaces_PRS_15(BundleID1) + "', '" + Insertion_Espaces_PRS_15(BundleID2) + "', '" + Insertion_Espaces_PRS_10(OPGID3) + "', '" + Insertion_Espaces_PRS_15(BundleID3) + "', '" + Insertion_Espaces_PRS_10(OPGID4) + "', '" + Insertion_Espaces_PRS_15(BundleID4) + "', '" + Insertion_Espaces_PRS_10(OPGID5) + "', '" + Insertion_Espaces_PRS_15(BundleID5) + "', '" + Insertion_Espaces_PRS_10(OPGID6) + "', '" + Insertion_Espaces_PRS_15(BundleID6) + "', '" + Insertion_Espaces_PRS_17(R0[r]) + "', '" + Insertion_Espaces_PRS_30(HPInvoiceNo) + "', '" + Insertion_Espaces_PRS_17(EndUserID) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3(PartPurchPriceCC) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_100(PartnerComment) + "', '" + Insertion_Espaces_PRS_15(PartnerReportedCBN) + "', '" + Insertion_Espaces_PRS_30(PartnerReference) + "', '" + Insertion_Espaces_PRS_40(PartnerInterTransID) + "', '" + Insertion_Espaces_PRS_6(IsDropShip) + "', '" + Insertion_Espaces_PRS_30(CustChannelPurchId) + "', '" + Insertion_Espaces_PRS_15(PurchaseAgreement) + "', '" + Insertion_Espaces_PRS_30(ReporterPurchOrderID) + "', '" + Insertion_Espaces_PRS_50(SuppliesTrackID) + "', '" + Insertion_Espaces_PRS_10(IntercompanyFlag) + "', '" + Insertion_Espaces_PRS_1600(ProdSerialIdHP) + "')";

                                    SqlCommand Cmd_Remp_PRS = new SqlCommand(Query_INS_PRS, Con_HP);
                                    Cmd_Remp_PRS.ExecuteNonQuery();

                                }

                                Trouve = true;
                                Reader_RS.Close();

                            }

                        }
                        Reader_Comp_KIT.Close();
                        if (Trouve == true)
                        {
                            String Query_Delete = "delete from [" + textBox5.Text + "PRS_I]  where [HPProductNumber] = '" + R2[r] + "'and [SoldToCustomerID] = '" + R0[r] + "'";
                            SqlCommand Cmd_Del = new SqlCommand(Query_Delete, Con_HP);
                            Cmd_Del.ExecuteNonQuery();

                        }
                        r++;
                    }
                    Con_KIT.Close();
                    Con_HP.Close();
                }

                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                SqlConnection con4 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                con4.Open();
                string myquery2 = " insert into [Incidents_HP] ([Numero Etat HP] ,[Incident]) values (" + textBox5.Text.ToString() + ",'" + e.Message.ToString().Replace("'", "") + "')";

                SqlCommand Incident = new SqlCommand(myquery2, con4);
                Incident.ExecuteNonQuery();
                con4.Close();
                if (codeHP == CodeHPE)
                {
                    /*Notification_Mail*/
                    SendEmail("notifhp@disway.com", "notifhp@disway.com", "Alerte HPE", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de PRS_E: " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                else
                {
                    /*Notification_Mail*/
                    SendEmail("notifhp@disway.com", "notifhp@disway.com", "Alerte HPI", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de PRS_I: " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                return false;
            }
        }

        /* Fonction de création de la table CUS_E: CUSTOMERS*/
        private bool Creation_CUS_Table(string codeHP)
        {
            try
            {
                if (codeHP == CodeHPE)
                {
                    SqlConnection con5_CUS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                    con5_CUS.Open();
                    string myquery = "CREATE TABLE [" + textBox5.Text + "CUS_E] (SoldToOrInvoicetoCustomerID Varchar (50) ,ReservedText  Varchar (50) ,CustomerName Varchar (50) ,Street  Varchar (MAX) ,PostCode  Varchar (50) ,CityName  Varchar (50) ,CountrySubEntityCode Varchar (50),CoutryCode Varchar (50) ,TaxIdentifier  Varchar (50),Street2  Varchar (MAX))";
                    SqlCommand Cmd_HDR =
                        new SqlCommand(myquery, con5_CUS);
                    Cmd_HDR.ExecuteNonQuery();
                    con5_CUS.Close();
                    /* MessageBox.Show("OK CUS_E");*/
                    richTextBox3.Text = richTextBox3.Text + " / " + "OK CUS_E";
                }
                else
                {
                    SqlConnection con5_CUS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                    con5_CUS.Open();
                    string myquery = "CREATE TABLE [" + textBox5.Text + "CUS_I] (SoldToOrInvoicetoCustomerID Varchar (50) ,ReservedText  Varchar (50) ,CustomerName Varchar (50) ,Street  Varchar (MAX) ,PostCode  Varchar (50) ,CityName  Varchar (50) ,CountrySubEntityCode Varchar (50),CoutryCode Varchar (50) ,TaxIdentifier  Varchar (50),Street2  Varchar (MAX))";
                    SqlCommand Cmd_HDR =
                        new SqlCommand(myquery, con5_CUS);
                    Cmd_HDR.ExecuteNonQuery();
                    con5_CUS.Close();
                    /* MessageBox.Show("OK CUS_E");*/
                    richTextBox3.Text = richTextBox3.Text + " / " + "OK CUS_I";
                }
                return true;
            }
            catch(Exception e )
            {
                MessageBox.Show(e.Message);
                SqlConnection con4 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                con4.Open();
                string myquery2 = " insert into [Incidents_HP] ([Numero Etat HP] ,[Incident]) values (" + textBox5.Text.ToString() + ",'" + e.Message.ToString().Replace("'", "") + "')"; 
                SqlCommand Incident = new SqlCommand(myquery2, con4);
                Incident.ExecuteNonQuery();
                con4.Close();
                if (codeHP == CodeHPE)
                {
                    /*Notification_Mail*/
                    SendEmail(" notifhp@disway.com", "notifhp@disway.com", "Alerte HPE", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de CUS_E : " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                else
                {
                    /*Notification_Mail*/
                    SendEmail(" notifhp@disway.com", "notifhp@disway.com", "Alerte HPI", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de CUS_I : " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                return false;
            }
        }

        /* Fonction de remplissage de la table CUS_E */
        private bool Remplissage_CUS(string codeHP)
        {
            try
            {
                if (codeHP == CodeHPE)
                {
                    /* Selection des codes familles HPE */
                    SqlConnection Con_FHP = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
                    Con_FHP.Open();
                    String Query = "select DISTINCT [code_Famille] from Famille_HP_New where Cartes='HPE'";
                    String Query2 = "Select count(*) from Famille_HP_New where Cartes='HPE'";
                    SqlCommand cmd_FamilleHP = new SqlCommand(Query, Con_FHP);
                    SqlCommand cmd_Count_FamilleHP = new SqlCommand(Query2, Con_FHP);
                    int Count_HP = (int)cmd_Count_FamilleHP.ExecuteScalar();
                    SqlDataReader Reader = cmd_FamilleHP.ExecuteReader();
                    string[] T = new string[Count_HP];
                    string H = "";
                    int i = 0;
                    while (Reader.Read())
                    {
                        T[i] = Reader.GetString(0);
                        i++;
                    }
                    Reader.Close();
                    for (int j = 0; j < Count_HP; j++)
                    {
                        if (j < Count_HP - 1)
                            H += "'" + T[j] + "',";
                        else
                            H += "'" + T[j] + "'";


                    }
                    Con_FHP.Close();

                    /* Selection des Clients exclus */

                    Con_FHP.Open();
                    String Query_cus = "select [Code_Client_Exclu] from Clients_Exclus";
                    String Querycus_cus_count = "Select count(*) from Clients_Exclus";
                    SqlCommand cmd_Cus_Exclus = new SqlCommand(Query_cus, Con_FHP);
                    SqlCommand cmd_Count_Clients_Exclus = new SqlCommand(Querycus_cus_count, Con_FHP);
                    int Count_Cus_Exclus = (int)cmd_Count_Clients_Exclus.ExecuteScalar();
                    SqlDataReader Reader_Cus = cmd_Cus_Exclus.ExecuteReader();
                    string[] R = new string[Count_Cus_Exclus];
                    string L = "";
                    int m = 0;
                    while (Reader_Cus.Read())
                    {
                        R[m] = Reader_Cus.GetString(0);
                        m++;
                    }
                    Reader_Cus.Close();
                    for (int n = 0; n < Count_Cus_Exclus; n++)
                    {
                        if (n < Count_Cus_Exclus - 1)
                            L += "'" + R[n] + "',";
                        else
                            L += "'" + R[n] + "'";
                    }
                    string CLT = "";
                    if (L != "") CLT = "AND (S.[N° donneur d'ordre] NOT IN (" + L + ")) ";
                    Con_FHP.Close();

                    /* Selection des Produits exclus */
                    Con_FHP.Open();
                    String Query_Prod = "select  [Code_Produit_Exclu] from Produits_Exclus";
                    String Querycus_Prod_Count = "Select count(*) from Produits_Exclus";
                    SqlCommand cmdProd_Exclus = new SqlCommand(Query_Prod, Con_FHP);
                    SqlCommand cmd_Count_Prod_Exclus = new SqlCommand(Querycus_Prod_Count, Con_FHP);
                    cmdProd_Exclus.CommandTimeout = 90;
                    int Count_Prod_Exclus = (int)cmd_Count_Prod_Exclus.ExecuteScalar();
                    SqlDataReader Reader_Prod = cmdProd_Exclus.ExecuteReader();
                    string[] P = new string[Count_Prod_Exclus];
                    string M = "";
                    int s = 0;
                    while (Reader_Prod.Read())
                    {
                        P[s] = Reader_Prod.GetString(0);
                        s++;
                    }
                    Reader_Prod.Close();
                    for (int r = 0; r < Count_Prod_Exclus; r++)
                    {
                        if (r < Count_Prod_Exclus - 1)
                            M += "'" + P[r] + "',";
                        else
                            M += "'" + P[r] + "'";


                    }
                    string PRD = "";
                    if (M != "") PRD = "AND (P.[N°] NOT IN (" + M + ")) ";
                    Con_FHP.Close();

                    /* Selection des informations de remplissage CUS_E */
                    SqlConnection con6_CUS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Disway;Integrated Security=True");
                    con6_CUS.Open();
                    string myquery =
    "SELECT DISTINCT S.[N° donneur d'ordre], C.Nom, C.Adresse, C.[Code postal], C.Ville,C.[Code pays] FROM dbo.___Sold_CM_BL S INNER JOIN dbo.___Product P  ON S.[N°] = P.[N°] INNER JOIN dbo.___Customer C ON S.[N° donneur d'ordre] = C.[N°] WHERE (P.[Code Famille] IN (" + H + ") " + Get_CodeSFamille(codeHP) + " AND  (S.[Date comptabilisation] BETWEEN '" + dateTimePicker1.Text + "' AND '" + dateTimePicker2.Text + "') " + CLT + PRD;

                    if (L == "") /* Si la liste des clients exclus est vide on enleve la condition sur les clients exclus */
                    {
                        string myquery_Light =
    "SELECT DISTINCT S.[N° donneur d'ordre], C.Nom, C.Adresse, C.[Code postal], C.Ville,C.[Code pays] FROM dbo.___Sold_CM_BL S INNER JOIN dbo.___Product P  ON S.[N°] = P.[N°] INNER JOIN dbo.___Customer C ON S.[N° donneur d'ordre] = C.[N°] WHERE (P.[Code Famille] IN (" + H + ") " + Get_CodeSFamille(codeHP) + " AND  (S.[Date comptabilisation] BETWEEN '" + dateTimePicker1.Text + "' AND '" + dateTimePicker2.Text + "') " + PRD;
                        /*    MessageBox.Show("Exécution de la requête sans clients exclus en cours ... !"); */

                        SqlCommand Cmd_CUS = new SqlCommand(myquery_Light, con6_CUS);
                        Cmd_CUS.CommandTimeout = 90;
                        SqlDataReader Reader_CUS = Cmd_CUS.ExecuteReader();
                        /* MessageBox.Show(myquery_Light); */
                        /* On récupère le nombre de lignes reourné par la requête*/
                        DataTable CUS_DT = new DataTable();
                        CUS_DT.Load(Reader_CUS);
                        int Count_myquery_Light = CUS_DT.Rows.Count;
                        /* MessageBox.Show(" Le nombre de lignes retourné est : " + Count_myquery_Light.ToString()  + ""); */
                        Reader_CUS = Cmd_CUS.ExecuteReader();
                        /*  */
                        int Count = Count_myquery_Light + 1;

                        string[] Z = new string[Count];
                        string[] Y = new string[Count];
                        string[] G = new string[Count];
                        string[] F = new string[Count];
                        string[] J = new string[Count];
                        string[] K = new string[Count];
                        int b = 0;
                        /* On récupère chaque ligne dans des tableaux */
                        while (Reader_CUS.Read())
                        {
                            Z[b] = Reader_CUS.GetString(0);
                            Y[b] = Reader_CUS.GetString(1);
                            G[b] = Reader_CUS.GetString(2);
                            F[b] = Reader_CUS.GetString(3);
                            J[b] = Reader_CUS.GetString(4);
                            K[b] = Reader_CUS.GetString(5);
                            b++;
                        }
                        Reader_CUS.Close();
                        /* Insertion des valeurs selectionnées */
                        for (int j = 0; j < Count - 1; j++)
                        {
                            SqlConnection con_Ins_CUS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                            con_Ins_CUS.Open();

                            string customer_Name = Y[j].Replace("'", "");
                            if (customer_Name.Length >= 35)
                            {
                                customer_Name = customer_Name.Substring(0, 34);
                            }
                            string customer_street = G[j].Replace("'", "");
                            if (customer_street.Length >= 35)
                            {
                                customer_street = customer_street.Substring(0, 35);
                            }

                            if (customer_street.Length == 0)
                            {
                                customer_street = " Rue "; //Champs obligatoire: on le remplit pour éviter le blocage s'il n'est pas renseigné
                            }
                            string customer_street2 = " ";                                                        
                            //String Query_INS_CUS = " insert into  [" + textBox5.Text + "CUS_E]  ([SoldToOrInvoicetoCustomerID],[ReservedText],[CustomerName],[Street], [PostCode],[CityName], [CountrySubEntityCode],[CoutryCode],[TaxIdentifier] ) VALUES ('" + Insertion_Espaces_AMS1(Z[j]) + "','" + Insertion_Espaces_CUS7("") + "', '" + Insertion_Espaces_CUS_35(customer_Name).Substring(0, 35) + "', '" + Insertion_Espaces_CUS_35(customer_street).Substring(0, 35) + "', '" + Insertion_Espaces_CUS_9(F[j].Replace(" ", "")).Substring(0, 9) + "','" + Insertion_Espaces_CUS_35(J[j]).Substring(0, 35) + "' , '" + Insertion_Espaces_CUS_9(K[j]).Substring(0, 9) + "', '" + Insertion_Espaces_CUS_2(K[j]) + "','" + Insertion_Espaces_CUS_20("").Substring(0, 20) + "')";
                            String Query_INS_CUS = " insert into  [" + textBox5.Text + "CUS_E]  ([SoldToOrInvoicetoCustomerID],[ReservedText],[CustomerName],[Street], [PostCode],[CityName], [CountrySubEntityCode],[CoutryCode],[TaxIdentifier],[Street2]) VALUES ('" + Insertion_Espaces_AMS1(Z[j]) + "','" + Insertion_Espaces_CUS7("") + "', '" + Insertion_Espaces_CUS_35(customer_Name).Substring(0, 35) + "', '" + Insertion_Espaces_CUS_35(customer_street).Substring(0, 35) + "', '" + Insertion_Espaces_CUS_9(F[j].Replace(" ", "")).Substring(0, 9) + "','" + Insertion_Espaces_CUS_35(J[j]).Substring(0, 35) + "' , '" + Insertion_Espaces_CUS_9(K[j]).Substring(0, 9) + "', '" + Insertion_Espaces_CUS_2(K[j]) + "','" + Insertion_Espaces_CUS_20("").Substring(0, 20) + "', '" + Insertion_Espaces_CUS_35(customer_street2).Substring(0, 35) + "')";
                            
                            /*  MessageBox.Show(Query_INS_CUS); */
                            SqlCommand Cmd_Remp_CUS = new SqlCommand(Query_INS_CUS, con_Ins_CUS);
                            Cmd_Remp_CUS.ExecuteNonQuery();
                            con_Ins_CUS.Close();
                            /* }
                             */
                        }


                    }
                    else /* Si on a des clients exclus on exéccute la requête totale */
                    {
                        /*   MessageBox.Show("Exécution de la requête totale en cours ... ! "); */
                        SqlCommand Cmd_HDR = new SqlCommand(myquery, con6_CUS);
                        Cmd_HDR.CommandTimeout = 90;
                        SqlDataReader Reader_CUS = Cmd_HDR.ExecuteReader();
                        /* */
                        DataTable CUS_DT = new DataTable();
                        CUS_DT.Load(Reader_CUS);
                        int Count_myquery_Light = CUS_DT.Rows.Count;
                        /*  MessageBox.Show(" Le nombre de lignes retourné est : " + Count_myquery_Light.ToString() + ""); */
                        Reader_CUS = Cmd_HDR.ExecuteReader();
                        /*  */
                        int Count = Count_myquery_Light + 1;

                        string[] Z = new string[Count];
                        string[] Y = new string[Count];
                        string[] G = new string[Count];
                        string[] F = new string[Count];
                        string[] J = new string[Count];
                        string[] K = new string[Count];
                        int b = 0;
                        /* On récupère chaque ligne dans des tableaux */
                        while (Reader_CUS.Read())
                        {
                            Z[b] = Reader_CUS.GetString(0);
                            Y[b] = Reader_CUS.GetString(1);
                            G[b] = Reader_CUS.GetString(2);
                            F[b] = Reader_CUS.GetString(3);
                            J[b] = Reader_CUS.GetString(4);
                            K[b] = Reader_CUS.GetString(5);
                            b++;
                        }
                        Reader_CUS.Close();
                        for (int j = 0; j < Count - 1; j++)
                        {
                            SqlConnection con_Ins_CUS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                            con_Ins_CUS.Open();
                            /*string customer_Name = Y[j].Replace('\'', ' ');*/
                            //string customer_Name = Y[j].Replace("'", "");
                            //string customer_street = G[j].Replace("'", "");
                            string customer_Name = Y[j].Replace("'", "");
                            if (customer_Name.Length >= 35)
                            {
                                customer_Name = customer_Name.Substring(0, 34);
                            }
                            string customer_street = G[j].Replace("'", "");
                            if (customer_street.Length >= 35)
                            {
                                customer_street = customer_street.Substring(0, 35);
                            }

                            if (customer_street.Length == 0)
                            {
                                customer_street = " Rue "; //Champs obligatoire: on le remplit pour éviter le blocage s'il n'est pas renseigné
                            }
                            string customer_street2 = " ";
                            //String Query_INS_CUS = " insert into  [" + textBox5.Text + "CUS_E]  ([SoldToOrInvoicetoCustomerID],[ReservedText],[CustomerName],[Street], [PostCode],[CityName], [CountrySubEntityCode],[CoutryCode], [TaxIdentifier] ) VALUES ('" + Insertion_Espaces_AMS1(Z[j]) + "','" + Insertion_Espaces_CUS7("") + "', '" + Insertion_Espaces_CUS_35(customer_Name) + "', '" + Insertion_Espaces_CUS_35(customer_street) + "', '" + Insertion_Espaces_CUS_9(F[j]) + "','" + Insertion_Espaces_CUS_35(J[j]) + "' , '" + Insertion_Espaces_CUS_9("MA") + "', '" + Insertion_Espaces_CUS_2(K[j]) + "','" + Insertion_Espaces_CUS_20("") + "')";
                            //String Query_INS_CUS = " insert into  [" + textBox5.Text + "CUS_E]  ([SoldToOrInvoicetoCustomerID],[ReservedText],[CustomerName],[Street], [PostCode],[CityName], [CountrySubEntityCode],[CoutryCode],[TaxIdentifier] ) VALUES ('" + Insertion_Espaces_AMS1(Z[j]) + "','" + Insertion_Espaces_CUS7("") + "', '" + Insertion_Espaces_CUS_35(customer_Name).Substring(0, 35) + "', '" + Insertion_Espaces_CUS_35(customer_street).Substring(0, 35) + "', '" + Insertion_Espaces_CUS_9(F[j].Replace(" ", "")).Substring(0, 9) + "','" + Insertion_Espaces_CUS_35(J[j]).Substring(0, 35) + "' , '" + Insertion_Espaces_CUS_9(K[j]).Substring(0, 9) + "', '" + Insertion_Espaces_CUS_2(K[j]) + "','" + Insertion_Espaces_CUS_20("").Substring(0, 20) + "')";
                            String Query_INS_CUS = " insert into  [" + textBox5.Text + "CUS_E]  ([SoldToOrInvoicetoCustomerID],[ReservedText],[CustomerName],[Street], [PostCode],[CityName], [CountrySubEntityCode],[CoutryCode],[TaxIdentifier],[Street2]) VALUES ('" + Insertion_Espaces_AMS1(Z[j]) + "','" + Insertion_Espaces_CUS7("") + "', '" + Insertion_Espaces_CUS_35(customer_Name).Substring(0, 35) + "', '" + Insertion_Espaces_CUS_35(customer_street).Substring(0, 35) + "', '" + Insertion_Espaces_CUS_9(F[j].Replace(" ", "")).Substring(0, 9) + "','" + Insertion_Espaces_CUS_35(J[j]).Substring(0, 35) + "' , '" + Insertion_Espaces_CUS_9(K[j]).Substring(0, 9) + "', '" + Insertion_Espaces_CUS_2(K[j]) + "','" + Insertion_Espaces_CUS_20("").Substring(0, 20) + "', '" + Insertion_Espaces_CUS_35(customer_street2).Substring(0, 35) + "')";

                            /*  MessageBox.Show(Query_INS_CUS); */
                            SqlCommand Cmd_Remp_CUS = new SqlCommand(Query_INS_CUS, con_Ins_CUS);
                            Cmd_Remp_CUS.ExecuteNonQuery();
                            con_Ins_CUS.Close();
                        }

                    }
                    con6_CUS.Close();
                    richTextBox4.Text = richTextBox4.Text + "  /  " + "OK CUS_E";
                }
                else
                {
                    /* Selection des codes familles HPI */
                    SqlConnection Con_FHP = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
                    Con_FHP.Open();
                    String Query = "select DISTINCT [code_Famille] from Famille_HP_New where Cartes='HPI'";
                    String Query2 = "Select count(*) from Famille_HP_New where Cartes='HPI'";
                    SqlCommand cmd_FamilleHP = new SqlCommand(Query, Con_FHP);
                    SqlCommand cmd_Count_FamilleHP = new SqlCommand(Query2, Con_FHP);
                    int Count_HP = (int)cmd_Count_FamilleHP.ExecuteScalar();
                    SqlDataReader Reader = cmd_FamilleHP.ExecuteReader();
                    string[] T = new string[Count_HP];
                    string H = "";
                    int i = 0;
                    while (Reader.Read())
                    {
                        T[i] = Reader.GetString(0);
                        i++;
                    }
                    Reader.Close();
                    for (int j = 0; j < Count_HP; j++)
                    {
                        if (j < Count_HP - 1)
                            H += "'" + T[j] + "',";
                        else
                            H += "'" + T[j] + "'";


                    }
                    Con_FHP.Close();

                    /* Selection des Clients exclus */

                    Con_FHP.Open();
                    String Query_cus = "select [Code_Client_Exclu] from Clients_Exclus";
                    String Querycus_cus_count = "Select count(*) from Clients_Exclus";
                    SqlCommand cmd_Cus_Exclus = new SqlCommand(Query_cus, Con_FHP);
                    SqlCommand cmd_Count_Clients_Exclus = new SqlCommand(Querycus_cus_count, Con_FHP);
                    int Count_Cus_Exclus = (int)cmd_Count_Clients_Exclus.ExecuteScalar();
                    SqlDataReader Reader_Cus = cmd_Cus_Exclus.ExecuteReader();
                    string[] R = new string[Count_Cus_Exclus];
                    string L = "";
                    int m = 0;
                    while (Reader_Cus.Read())
                    {
                        R[m] = Reader_Cus.GetString(0);
                        m++;
                    }
                    Reader_Cus.Close();
                    for (int n = 0; n < Count_Cus_Exclus; n++)
                    {
                        if (n < Count_Cus_Exclus - 1)
                            L += "'" + R[n] + "',";
                        else
                            L += "'" + R[n] + "'";
                    }
                    string CLT = "";
                    if (L != "") CLT = "AND (S.[N° donneur d'ordre] NOT IN (" + L + ")) ";
                    Con_FHP.Close();

                    /* Selection des Produits exclus */
                    Con_FHP.Open();
                    String Query_Prod = "select  [Code_Produit_Exclu] from Produits_Exclus";
                    String Querycus_Prod_Count = "Select count(*) from Produits_Exclus";
                    SqlCommand cmdProd_Exclus = new SqlCommand(Query_Prod, Con_FHP);
                    SqlCommand cmd_Count_Prod_Exclus = new SqlCommand(Querycus_Prod_Count, Con_FHP);
                    cmdProd_Exclus.CommandTimeout = 90;
                    int Count_Prod_Exclus = (int)cmd_Count_Prod_Exclus.ExecuteScalar();
                    SqlDataReader Reader_Prod = cmdProd_Exclus.ExecuteReader();
                    string[] P = new string[Count_Prod_Exclus];
                    string M = "";
                    int s = 0;
                    while (Reader_Prod.Read())
                    {
                        P[s] = Reader_Prod.GetString(0);
                        s++;
                    }
                    Reader_Prod.Close();
                    for (int r = 0; r < Count_Prod_Exclus; r++)
                    {
                        if (r < Count_Prod_Exclus - 1)
                            M += "'" + P[r] + "',";
                        else
                            M += "'" + P[r] + "'";


                    }
                    string PRD = "";
                    if (M != "") PRD = "AND (P.[N°] NOT IN (" + M + ")) ";
                    Con_FHP.Close();

                    /* Selection des informations de remplissage CUS_E */
                    SqlConnection con6_CUS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Disway;Integrated Security=True");
                    con6_CUS.Open();
                    string myquery =
    "SELECT DISTINCT S.[N° donneur d'ordre], C.Nom, C.Adresse, C.[Code postal], C.Ville,C.[Code pays] FROM dbo.___Sold_CM_BL S INNER JOIN dbo.___Product P  ON S.[N°] = P.[N°] INNER JOIN dbo.___Customer C ON S.[N° donneur d'ordre] = C.[N°] WHERE (P.[Code Famille] IN (" + H + ")) " + Get_CodeSFamille(codeHP) + " AND  (S.[Date comptabilisation] BETWEEN '" + dateTimePicker1.Text + "' AND '" + dateTimePicker2.Text + "') " + PRD + CLT;

                    if (L == "") /* Si la liste des clients exclus est vide on enleve la condition sur les clients exclus */
                    {
                        string myquery_Light =
    "SELECT DISTINCT S.[N° donneur d'ordre], C.Nom, C.Adresse, C.[Code postal], C.Ville,C.[Code pays] FROM dbo.___Sold_CM_BL S INNER JOIN dbo.___Product P  ON S.[N°] = P.[N°] INNER JOIN dbo.___Customer C ON S.[N° donneur d'ordre] = C.[N°] WHERE (P.[Code Famille] IN (" + H + ")) " + Get_CodeSFamille(codeHP) + " AND  (S.[Date comptabilisation] BETWEEN '" + dateTimePicker1.Text + "' AND '" + dateTimePicker2.Text + "') " + PRD;
                        /*    MessageBox.Show("Exécution de la requête sans clients exclus en cours ... !"); */

                        SqlCommand Cmd_CUS = new SqlCommand(myquery_Light, con6_CUS);
                        Cmd_CUS.CommandTimeout = 90;
                        SqlDataReader Reader_CUS = Cmd_CUS.ExecuteReader();
                        /* MessageBox.Show(myquery_Light); */
                        /* On récupère le nombre de lignes reourné par la requête*/
                        DataTable CUS_DT = new DataTable();
                        CUS_DT.Load(Reader_CUS);
                        int Count_myquery_Light = CUS_DT.Rows.Count;
                        /* MessageBox.Show(" Le nombre de lignes retourné est : " + Count_myquery_Light.ToString()  + ""); */
                        Reader_CUS = Cmd_CUS.ExecuteReader();
                        /*  */
                        int Count = Count_myquery_Light + 1;

                        string[] Z = new string[Count];
                        string[] Y = new string[Count];
                        string[] G = new string[Count];
                        string[] F = new string[Count];
                        string[] J = new string[Count];
                        string[] K = new string[Count];
                        int b = 0;
                        /* On récupère chaque ligne dans des tableaux */
                        while (Reader_CUS.Read())
                        {
                            Z[b] = Reader_CUS.GetString(0);
                            Y[b] = Reader_CUS.GetString(1);
                            G[b] = Reader_CUS.GetString(2);
                            F[b] = Reader_CUS.GetString(3);
                            J[b] = Reader_CUS.GetString(4);
                            K[b] = Reader_CUS.GetString(5);
                            b++;
                        }
                        Reader_CUS.Close();
                        /* Insertion des valeurs selectionnées */
                        for (int j = 0; j < Count - 1; j++)
                        {
                            SqlConnection con_Ins_CUS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                            con_Ins_CUS.Open();

                            string customer_Name = Y[j].Replace("'", "");
                            if (customer_Name.Length >= 35)
                            {
                                customer_Name = customer_Name.Substring(0, 34);
                            }
                            string customer_street = G[j].Replace("'", "");
                            if (customer_street.Length >= 35)
                            {
                                customer_street = customer_street.Substring(0, 35);
                            }
                            
                            /* -------- AURORA */
                            if (customer_street.Length == 0)
                            {
                                customer_street = " Rue "; //Champs obligatoire: on le remplit pour éviter le blocage s'il n'est pas renseigné
                            }
                            string customer_street2 = " ";
                            //String Query_INS_CUS = " insert into  [" + textBox5.Text + "CUS_I]  ([SoldToOrInvoicetoCustomerID],[ReservedText],[CustomerName],[Street], [PostCode],[CityName], [CountrySubEntityCode],[CoutryCode],[TaxIdentifier] ) VALUES ('" + Insertion_Espaces_AMS1(Z[j]) + "','" + Insertion_Espaces_CUS7("") + "', '" + Insertion_Espaces_CUS_35(customer_Name).Substring(0, 35) + "', '" + Insertion_Espaces_CUS_35(customer_street).Substring(0, 35) + "', '" + Insertion_Espaces_CUS_9(F[j].Replace(" ", "")).Substring(0, 9) + "','" + Insertion_Espaces_CUS_35(J[j]).Substring(0, 35) + "' , '" + Insertion_Espaces_CUS_9(K[j]).Substring(0, 9) + "', '" + Insertion_Espaces_CUS_2(K[j]) + "','" + Insertion_Espaces_CUS_20("").Substring(0, 20) + "')";
                            String Query_INS_CUS = " insert into  [" + textBox5.Text + "CUS_I]  ([SoldToOrInvoicetoCustomerID],[ReservedText],[CustomerName],[Street], [PostCode],[CityName], [CountrySubEntityCode],[CoutryCode],[TaxIdentifier],[Street2] ) VALUES ('" + Insertion_Espaces_AMS1(Z[j]) + "','" + Insertion_Espaces_CUS7("") + "', '" + Insertion_Espaces_CUS_35(customer_Name).Substring(0, 35) + "', '" + Insertion_Espaces_CUS_35(customer_street).Substring(0, 35) + "', '" + Insertion_Espaces_CUS_9(F[j].Replace(" ", "")).Substring(0, 9) + "','" + Insertion_Espaces_CUS_35(J[j]).Substring(0, 35) + "' , '" + Insertion_Espaces_CUS_9(K[j]).Substring(0, 9) + "', '" + Insertion_Espaces_CUS_2(K[j]) + "','" + Insertion_Espaces_CUS_20("").Substring(0, 20) + "', '" + Insertion_Espaces_CUS_35(customer_street2).Substring(0, 35) + "')";
                            /* -------- AURORA */

                            /*  MessageBox.Show(Query_INS_CUS); */
                            SqlCommand Cmd_Remp_CUS = new SqlCommand(Query_INS_CUS, con_Ins_CUS);
                            Cmd_Remp_CUS.ExecuteNonQuery();
                            con_Ins_CUS.Close();
                            /* }
                             */
                        }


                    }
                    else /* Si on a des clients exclus on exéccute la requête totale */
                    {
                        /*   MessageBox.Show("Exécution de la requête totale en cours ... ! "); */
                        SqlCommand Cmd_HDR = new SqlCommand(myquery, con6_CUS);
                        Cmd_HDR.CommandTimeout = 90;
                        SqlDataReader Reader_CUS = Cmd_HDR.ExecuteReader();
                        /* */
                        DataTable CUS_DT = new DataTable();
                        CUS_DT.Load(Reader_CUS);
                        int Count_myquery_Light = CUS_DT.Rows.Count;
                        /*  MessageBox.Show(" Le nombre de lignes retourné est : " + Count_myquery_Light.ToString() + ""); */
                        Reader_CUS = Cmd_HDR.ExecuteReader();
                        /*  */
                        int Count = Count_myquery_Light + 1;

                        string[] Z = new string[Count];
                        string[] Y = new string[Count];
                        string[] G = new string[Count];
                        string[] F = new string[Count];
                        string[] J = new string[Count];
                        string[] K = new string[Count];
                        int b = 0;
                        /* On récupère chaque ligne dans des tableaux */
                        while (Reader_CUS.Read())
                        {
                            Z[b] = Reader_CUS.GetString(0);
                            Y[b] = Reader_CUS.GetString(1);
                            G[b] = Reader_CUS.GetString(2);
                            F[b] = Reader_CUS.GetString(3);
                            J[b] = Reader_CUS.GetString(4);
                            K[b] = Reader_CUS.GetString(5);
                            b++;
                        }
                        Reader_CUS.Close();
                        for (int j = 0; j < Count - 1; j++)
                        {
                            SqlConnection con_Ins_CUS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                            con_Ins_CUS.Open();
                            /*string customer_Name = Y[j].Replace('\'', ' ');*
                            string customer_Name = Y[j].Replace("'", "");
                            ring customer_street = G[j].Replace("'", "");st*/
                            string customer_Name = Y[j].Replace("'", "");
                            if (customer_Name.Length >= 35)
                            {
                                customer_Name = customer_Name.Substring(0, 34);
                            }
                            string customer_street = G[j].Replace("'", "");
                            if (customer_street.Length >= 35)
                            {
                                customer_street = customer_street.Substring(0, 35);
                            }
                                                        
                            /* -------- AURORA */
                            if (customer_street.Length == 0)
                            {
                                customer_street = " Rue "; //Champs obligatoire: on le remplit pour éviter le blocage s'il n'est pas renseigné
                            }
                            string customer_street2 = " ";
                            //String Query_INS_CUS = " insert into  [" + textBox5.Text + "CUS_I]  ([SoldToOrInvoicetoCustomerID],[ReservedText],[CustomerName],[Street], [PostCode],[CityName], [CountrySubEntityCode],[CoutryCode], [TaxIdentifier] ) VALUES ('" + Insertion_Espaces_AMS1(Z[j]) + "','" + Insertion_Espaces_CUS7("") + "', '" + Insertion_Espaces_CUS_35(customer_Name) + "', '" + Insertion_Espaces_CUS_35(customer_street) + "', '" + Insertion_Espaces_CUS_9(F[j]) + "','" + Insertion_Espaces_CUS_35(J[j]) + "' , '" + Insertion_Espaces_CUS_9("MA") + "', '" + Insertion_Espaces_CUS_2(K[j]) + "','" + Insertion_Espaces_CUS_20("") + "')";                        /*  MessageBox.Show(Query_INS_CUS); */
                            //String Query_INS_CUS = " insert into  [" + textBox5.Text + "CUS_I]  ([SoldToOrInvoicetoCustomerID],[ReservedText],[CustomerName],[Street], [PostCode],[CityName], [CountrySubEntityCode],[CoutryCode],[TaxIdentifier] ) VALUES ('" + Insertion_Espaces_AMS1(Z[j]) + "','" + Insertion_Espaces_CUS7("") + "', '" + Insertion_Espaces_CUS_35(customer_Name).Substring(0, 35) + "', '" + Insertion_Espaces_CUS_35(customer_street).Substring(0, 35) + "', '" + Insertion_Espaces_CUS_9(F[j].Replace(" ", "")).Substring(0, 9) + "','" + Insertion_Espaces_CUS_35(J[j]).Substring(0, 35) + "' , '" + Insertion_Espaces_CUS_9(K[j]).Substring(0, 9) + "', '" + Insertion_Espaces_CUS_2(K[j]) + "','" + Insertion_Espaces_CUS_20("").Substring(0, 20) + "')";
                            String Query_INS_CUS = " insert into  [" + textBox5.Text + "CUS_I]  ([SoldToOrInvoicetoCustomerID],[ReservedText],[CustomerName],[Street], [PostCode],[CityName], [CountrySubEntityCode],[CoutryCode],[TaxIdentifier],[Street2] ) VALUES ('" + Insertion_Espaces_AMS1(Z[j]) + "','" + Insertion_Espaces_CUS7("") + "', '" + Insertion_Espaces_CUS_35(customer_Name).Substring(0, 35) + "', '" + Insertion_Espaces_CUS_35(customer_street).Substring(0, 35) + "', '" + Insertion_Espaces_CUS_9(F[j].Replace(" ", "")).Substring(0, 9) + "','" + Insertion_Espaces_CUS_35(J[j]).Substring(0, 35) + "' , '" + Insertion_Espaces_CUS_9(K[j]).Substring(0, 9) + "', '" + Insertion_Espaces_CUS_2(K[j]) + "','" + Insertion_Espaces_CUS_20("").Substring(0, 20) + "', '" + Insertion_Espaces_CUS_35(customer_street2).Substring(0, 35) + "')";
                            /* -------- AURORA */

                            SqlCommand Cmd_Remp_CUS = new SqlCommand(Query_INS_CUS, con_Ins_CUS);
                            Cmd_Remp_CUS.ExecuteNonQuery();
                            con_Ins_CUS.Close();
                        }

                    }
                    con6_CUS.Close();
                    richTextBox4.Text = richTextBox4.Text + "  /  " + "OK CUS_I";
                }
                return true;
            }

            catch(Exception e)
            {
                MessageBox.Show("Remplissage CUS_E NOT OK !");
                MessageBox.Show(e.Message);
                SqlConnection con4 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                con4.Open();
                string myquery2 = " insert into [Incidents_HP] ([Numero Etat HP] ,[Incident]) values (" + textBox5.Text.ToString() + ",'" + e.Message.ToString().Replace("'", "") + "')";
                SqlCommand Incident = new SqlCommand(myquery2, con4);
                Incident.ExecuteNonQuery();
                con4.Close();
                if (codeHP == CodeHPE)
                {
                    /*Notification_Mail*/
                    SendEmail(" notifhp@disway.com", "notifhp@disway.com", "Alerte HPE", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de CUS_E: " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                else
                {
                    /*Notification_Mail*/
                    SendEmail(" notifhp@disway.com", "notifhp@disway.com", "Alerte HPI", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de CUS_I: " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                return false;
            }

        }

        /* Fonction de création de la table AMS_E */
        private bool Creation_AMS_Table(string codeHP)
        {
            try
            {
                if (codeHP == CodeHPE)
                {
                    SqlConnection con5_AMS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                    con5_AMS.Open();
                    string myquery = "CREATE TABLE [" + textBox5.Text + "AMS_E] ( CustomerID Varchar (50) PRIMARY KEY,CurrencyCode Varchar (50) ,SelloutAmount  Varchar (50) ,SpecialPrincingReference  Varchar (50))";

                    SqlCommand Cmd_HDR =
                        new SqlCommand(myquery, con5_AMS);
                    Cmd_HDR.ExecuteNonQuery();
                    con5_AMS.Close();
                    /* MessageBox.Show("OK AMS_E"); */
                    richTextBox3.Text = richTextBox3.Text + " / " + "OK AMS_E";
                }
                else
                {
                    SqlConnection con5_AMS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                    con5_AMS.Open();
                    string myquery = "CREATE TABLE [" + textBox5.Text + "AMS_I] ( CustomerID Varchar (50) PRIMARY KEY,CurrencyCode Varchar (50) ,SelloutAmount  Varchar (50) ,SpecialPrincingReference  Varchar (50))";

                    SqlCommand Cmd_HDR =
                        new SqlCommand(myquery, con5_AMS);
                    Cmd_HDR.ExecuteNonQuery();
                    con5_AMS.Close();
                    /* MessageBox.Show("OK AMS_E"); */
                    richTextBox3.Text = richTextBox3.Text + " / " + "OK AMS_I";
                }
                return true;
            }
            catch(Exception e )
            {
                MessageBox.Show(e.Message);
                SqlConnection con4 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                con4.Open();
                string myquery2 = " insert into [Incidents_HP] ([Numero Etat HP] ,[Incident]) values (" + textBox5.Text.ToString() + ",'" + e.Message.ToString().Replace("'", "") + "')";
                SqlCommand Incident = new SqlCommand(myquery2, con4);
                Incident.ExecuteNonQuery();
                con4.Close();
                if (codeHP == CodeHPE)
                {
                    /*Notification_Mail*/
                    SendEmail("notifhp@disway.com", "notifhp@disway.com", "Alerte HPE", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de AMS_E: " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                else
                {
                    /*Notification_Mail*/
                    SendEmail("notifhp@disway.com", "notifhp@disway.com", "Alerte HPI", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de AMS_I: " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                return false;
            }
        }
        /*Fonction de remplissage de la fonction AMS_E */
        private bool Remplissage_AMS(string codeHP)
        {
            try
            {
                if (codeHP == CodeHPE)
                {
                    /* Selection des codes familles HPE */
                    SqlConnection Con_FHP = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
                    Con_FHP.Open();
                    String Query = "select DISTINCT [code_Famille] from Famille_HP_New where Cartes='" + codeHP +"'";
                    String Query2 = "Select count(*) from Famille_HP_New where Cartes='" + codeHP +"'";
                    SqlCommand cmd_FamilleHP = new SqlCommand(Query, Con_FHP);
                    SqlCommand cmd_Count_FamilleHP = new SqlCommand(Query2, Con_FHP);
                    int Count_HP = (int)cmd_Count_FamilleHP.ExecuteScalar();
                    SqlDataReader Reader = cmd_FamilleHP.ExecuteReader();
                    string[] T = new string[Count_HP];
                    string H = "";
                    int i = 0;
                    while (Reader.Read())
                    {
                        T[i] = Reader.GetString(0);
                        i++;
                    }
                    Reader.Close();
                    for (int j = 0; j < Count_HP; j++)
                    {
                        if (j < Count_HP - 1)
                            H += "'" + T[j] + "',";
                        else
                            H += "'" + T[j] + "'";


                    }
                    Con_FHP.Close();

                    /* Selection des Clients exclus */

                    Con_FHP.Open();
                    String Query_cus = "select [Code_Client_Exclu] from Clients_Exclus";
                    String Querycus_cus_count = "Select count(*) from Clients_Exclus";
                    SqlCommand cmd_Cus_Exclus = new SqlCommand(Query_cus, Con_FHP);
                    SqlCommand cmd_Count_Clients_Exclus = new SqlCommand(Querycus_cus_count, Con_FHP);
                    int Count_Cus_Exclus = (int)cmd_Count_Clients_Exclus.ExecuteScalar();
                    SqlDataReader Reader_Cus = cmd_Cus_Exclus.ExecuteReader();
                    string[] R = new string[Count_Cus_Exclus];
                    string L = "";
                    int m = 0;
                    while (Reader_Cus.Read())
                    {
                        T[m] = Reader_Cus.GetString(0);
                        m++;
                    }
                    Reader_Cus.Close();
                    for (int n = 0; n < Count_Cus_Exclus; n++)
                    {
                        if (n < Count_Cus_Exclus - 1)
                            L += "'" + T[n] + "',";
                        else
                            L += "'" + T[n] + "'";
                    }
                    string CLT = "";
                    if (L != "") CLT = "AND (S.[N° donneur d'ordre] NOT IN (" + L + ")) ";
                    Con_FHP.Close();
                    /* Selection des Produits exclus */
                    Con_FHP.Open();
                    String Query_Prod = "select  [Code_Produit_Exclu] from Produits_Exclus";
                    String Querycus_Prod_Count = "Select count(*) from Produits_Exclus";
                    SqlCommand cmdProd_Exclus = new SqlCommand(Query_Prod, Con_FHP);
                    SqlCommand cmd_Count_Prod_Exclus = new SqlCommand(Querycus_Prod_Count, Con_FHP);
                    int Count_Prod_Exclus = (int)cmd_Count_Prod_Exclus.ExecuteScalar();
                    cmdProd_Exclus.CommandTimeout = 90;
                    SqlDataReader Reader_Prod = cmdProd_Exclus.ExecuteReader();
                    string[] P = new string[Count_Prod_Exclus];
                    string M = "";
                    int s = 0;
                    while (Reader_Prod.Read())
                    {
                        P[s] = Reader_Prod.GetString(0);
                        s++;
                    }
                    Reader_Prod.Close();
                    for (int r = 0; r < Count_Prod_Exclus; r++)
                    {
                        if (r < Count_Prod_Exclus - 1)
                            M += "'" + P[r] + "',";
                        else
                            M += "'" + P[r] + "'";


                    }
                    string PRD = "";
                    if (M != "") PRD = "AND (P.[N°] NOT IN (" + M + ")) ";
                    Con_FHP.Close();


                    /* Selection des informations de remplissage AMS_E */
                    SqlConnection con6_AMS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Disway;Integrated Security=True");
                    con6_AMS.Open();
                    string myquery =
    "SELECT S.[N° donneur d'ordre], S.[Code devise], CONVERT(numeric(18,2) ,ISNULL(SUM(S.Quantité * S.PU) ,0) ) AS Amount, MAX(S.[N° lot]) AS [N° lot] FROM dbo.___Sold_CM_BL S INNER JOIN dbo.___Product P ON S.[N°] = P.[N°] WHERE (P.[Code Famille] IN (" + H + ") " + Get_CodeSFamille(codeHP) + " AND (S.[Date comptabilisation] BETWEEN '" + dateTimePicker1.Text + "' AND '" + dateTimePicker2.Text + "') " + CLT + PRD + " GROUP BY S.[N° donneur d'ordre], S.[Code devise] ORDER BY S.[N° donneur d'ordre]";

                    if (L == "") /* Si la liste des clients exclus est vide on enleve la condition sur les clients exclus */
                    {
                        string myquery_Light =
    "SELECT S.[N° donneur d'ordre], S.[Code devise], CONVERT(numeric(18,2) ,ISNULL(SUM(S.Quantité * S.PU) ,0) ) AS Amount, MAX(S.[N° lot]) AS [N° lot] FROM dbo.___Sold_CM_BL S INNER JOIN dbo.___Product P ON S.[N°] = P.[N°] WHERE (P.[Code Famille] IN (" + H + ") " + Get_CodeSFamille(codeHP) + " AND  (S.[Date comptabilisation] BETWEEN '" + dateTimePicker1.Text + "' AND '" + dateTimePicker2.Text + "') " + PRD + " GROUP BY S.[N° donneur d'ordre], S.[Code devise] ORDER BY S.[N° donneur d'ordre]";
                        /*    MessageBox.Show("Exécution de la requête sans clients exclus  en cours ... !"); */
                        SqlCommand Cmd_HDR = new SqlCommand(myquery_Light, con6_AMS);
                        Cmd_HDR.CommandTimeout = 90;
                        SqlDataReader Reader_AMS = Cmd_HDR.ExecuteReader();

                        /* MessageBox.Show(myquery_Light); */
                        /* On récupère le nombre de lignes reourné par la requête*/
                        DataTable AMAS_DT = new DataTable();
                        AMAS_DT.Load(Reader_AMS);
                        int Count_myquery_Light = AMAS_DT.Rows.Count;
                        /*   MessageBox.Show(" Le nombre de lignes retourné est : " + Count_myquery_Light.ToString()  + ""); */
                        Reader_AMS = Cmd_HDR.ExecuteReader();
                        /*  */
                        int Count = Count_myquery_Light + 1;

                        string[] Z = new string[Count];
                        string[] Y = new string[Count];
                        string[] G = new string[Count];
                        string[] F = new string[Count];
                        int b = 0;
                        /* On récupère chaque ligne dans des tableaux */
                        while (Reader_AMS.Read())
                        {
                            Z[b] = Reader_AMS.GetString(0);
                            Y[b] = Reader_AMS.GetString(1);
                            G[b] = Reader_AMS[2].ToString();
                            F[b] = Reader_AMS.GetString(3);
                            b++;
                        }
                        Reader_AMS.Close();
                        /* Insertion des valeurs selectionnées */
                        for (int j = 0; j < Count - 1; j++)
                        {
                            /*  MessageBox.Show(" Le client et ses infos sont : " + Z[j] + " , " + Y[j] + " , " + G[j] + ", " + F[j] +""); */
                            SqlConnection con_Ins_AMS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                            con_Ins_AMS.Open();
                            if (Z[j] != "" && Y[j] != "" && G[j] != "")
                            {
                                String Query_INS_AMS = " insert into [" + textBox5.Text + "AMS_E] ([CustomerID],[CurrencyCode],[SelloutAmount],[SpecialPrincingReference]) VALUES ('" + Insertion_Espaces_AMS1(Z[j]) + "', '" + Y[j] + "', '" + Insertion_Negatif(Insertion_Espaces_AMS1(Insertion_ZERO_Float(G[j].Replace(',', '.')))) + "', '" + Insertion_Espaces_AMS2(F[j]) + "' )";
                                SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_AMS, con_Ins_AMS);
                                Cmd_Remp_AMS.ExecuteNonQuery();
                                con_Ins_AMS.Close();
                            }

                        }
                    }
                    else /* Si on a des clients exclus on exéccute la requête totale */
                    {
                        SqlCommand Cmd_HDR = new SqlCommand(myquery, con6_AMS);
                        SqlDataReader Reader_AMS = Cmd_HDR.ExecuteReader();
                        /* */
                        DataTable AMAS_DT = new DataTable();
                        AMAS_DT.Load(Reader_AMS);
                        int Count_myquery_Light = AMAS_DT.Rows.Count;
                        Reader_AMS = Cmd_HDR.ExecuteReader();
                        /*  */
                        int Count = Count_myquery_Light + 1;
                        string[] Z = new string[Count];
                        string[] Y = new string[Count];
                        string[] G = new string[Count];
                        string[] F = new string[Count];
                        int b = 0;
                        while (Reader_AMS.Read())
                        {
                            Z[b] = Reader_AMS.GetString(0);
                            Y[b] = Reader_AMS.GetString(1);
                            G[b] = Reader_AMS[2].ToString();
                            F[b] = Reader_AMS.GetString(3);
                            b++;
                        }
                        Reader_AMS.Close();
                        for (int j = 0; j < Count - 1; j++)
                        {
                            SqlConnection con_Ins_AMS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                            con_Ins_AMS.Open();
                            String Query_INS_AMS = " insert into  [" + textBox5.Text + "AMS_E] ([CustomerID],[CurrencyCode],[SelloutAmount],[SpecialPrincingReference]) VALUES ('" + Insertion_Espaces_AMS1(Z[j]) + "', '" + Y[j] + "', '" + Insertion_Negatif(Insertion_Espaces_AMS1(Insertion_ZERO_Float(G[j].Replace(',', '.')))) + "', '" + Insertion_Espaces_AMS2(F[j]) + "' )";
                            SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_AMS, con_Ins_AMS);
                            Cmd_Remp_AMS.ExecuteNonQuery();
                            con_Ins_AMS.Close();
                        }

                    }
                    con6_AMS.Close();
                    richTextBox4.Text = richTextBox4.Text + "  /  " + "OK AMS_E";
                }
                else
                {
                    /* Selection des codes familles HPI */
                    SqlConnection Con_FHP = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
                    Con_FHP.Open();
                    String Query = "select DISTINCT [code_Famille] from Famille_HP_New where Cartes='" + codeHP +"'";
                    String Query2 = "Select count(*) from Famille_HP_New where Cartes='" + codeHP +"'";
                    SqlCommand cmd_FamilleHP = new SqlCommand(Query, Con_FHP);
                    SqlCommand cmd_Count_FamilleHP = new SqlCommand(Query2, Con_FHP);
                    int Count_HP = (int)cmd_Count_FamilleHP.ExecuteScalar();
                    SqlDataReader Reader = cmd_FamilleHP.ExecuteReader();
                    string[] T = new string[Count_HP];
                    string H = "";
                    int i = 0;
                    while (Reader.Read())
                    {
                        T[i] = Reader.GetString(0);
                        i++;
                    }
                    Reader.Close();
                    for (int j = 0; j < Count_HP; j++)
                    {
                        if (j < Count_HP - 1)
                            H += "'" + T[j] + "',";
                        else
                            H += "'" + T[j] + "'";


                    }
                    Con_FHP.Close();

                    /* Selection des Clients exclus */

                    Con_FHP.Open();
                    String Query_cus = "select [Code_Client_Exclu] from Clients_Exclus";
                    String Querycus_cus_count = "Select count(*) from Clients_Exclus";
                    SqlCommand cmd_Cus_Exclus = new SqlCommand(Query_cus, Con_FHP);
                    SqlCommand cmd_Count_Clients_Exclus = new SqlCommand(Querycus_cus_count, Con_FHP);
                    int Count_Cus_Exclus = (int)cmd_Count_Clients_Exclus.ExecuteScalar();
                    SqlDataReader Reader_Cus = cmd_Cus_Exclus.ExecuteReader();
                    string[] R = new string[Count_Cus_Exclus];
                    string L = "";
                    int m = 0;
                    while (Reader_Cus.Read())
                    {
                        T[m] = Reader_Cus.GetString(0);
                        m++;
                    }
                    Reader_Cus.Close();
                    for (int n = 0; n < Count_Cus_Exclus; n++)
                    {
                        if (n < Count_Cus_Exclus - 1)
                            L += "'" + T[n] + "',";
                        else
                            L += "'" + T[n] + "'";
                    }
                    string CLT = "";
                    if (L != "") CLT = "AND (S.[N° donneur d'ordre] NOT IN (" + L + ")) ";
                    Con_FHP.Close();
                    /* Selection des Produits exclus */
                    Con_FHP.Open();
                    String Query_Prod = "select  [Code_Produit_Exclu] from Produits_Exclus";
                    String Querycus_Prod_Count = "Select count(*) from Produits_Exclus";
                    SqlCommand cmdProd_Exclus = new SqlCommand(Query_Prod, Con_FHP);
                    SqlCommand cmd_Count_Prod_Exclus = new SqlCommand(Querycus_Prod_Count, Con_FHP);
                    int Count_Prod_Exclus = (int)cmd_Count_Prod_Exclus.ExecuteScalar();
                    cmdProd_Exclus.CommandTimeout = 90;
                    SqlDataReader Reader_Prod = cmdProd_Exclus.ExecuteReader();
                    string[] P = new string[Count_Prod_Exclus];
                    string M = "";
                    int s = 0;
                    while (Reader_Prod.Read())
                    {
                        P[s] = Reader_Prod.GetString(0);
                        s++;
                    }
                    Reader_Prod.Close();
                    for (int r = 0; r < Count_Prod_Exclus; r++)
                    {
                        if (r < Count_Prod_Exclus - 1)
                            M += "'" + P[r] + "',";
                        else
                            M += "'" + P[r] + "'";


                    }
                    string PRD = "";
                    if (M != "") PRD = "AND (P.[N°] NOT IN (" + M + ")) ";
                    Con_FHP.Close();


                    /* Selection des informations de remplissage AMS_I */
                    SqlConnection con6_AMS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Disway;Integrated Security=True");
                    con6_AMS.Open();
                    string myquery =
    "SELECT S.[N° donneur d'ordre], S.[Code devise], CONVERT(numeric(18,2) ,ISNULL(SUM(S.Quantité * S.PU) ,0) ) AS Amount, MAX(S.[N° lot]) AS [N° lot] FROM dbo.___Sold_CM_BL S INNER JOIN dbo.___Product P ON S.[N°] = P.[N°] WHERE (P.[Code Famille] IN (" + H + ")) " + Get_CodeSFamille(codeHP) + " AND (S.[Date comptabilisation] BETWEEN '" + dateTimePicker1.Text + "' AND '" + dateTimePicker2.Text + "') " + CLT + PRD + " GROUP BY S.[N° donneur d'ordre], S.[Code devise] ORDER BY S.[N° donneur d'ordre]";

                    if (L == "") /* Si la liste des clients exclus est vide on enleve la condition sur les clients exclus */
                    {
                        string myquery_Light =
    "SELECT S.[N° donneur d'ordre], S.[Code devise], CONVERT(numeric(18,2) ,ISNULL(SUM(S.Quantité * S.PU) ,0) ) AS Amount, MAX(S.[N° lot]) AS [N° lot] FROM dbo.___Sold_CM_BL S INNER JOIN dbo.___Product P ON S.[N°] = P.[N°] WHERE (P.[Code Famille] IN (" + H + ")) " + Get_CodeSFamille(codeHP) + " AND  (S.[Date comptabilisation] BETWEEN '" + dateTimePicker1.Text + "' AND '" + dateTimePicker2.Text + "') " + PRD + " GROUP BY S.[N° donneur d'ordre], S.[Code devise] ORDER BY S.[N° donneur d'ordre]";
                        /*    MessageBox.Show("Exécution de la requête sans clients exclus  en cours ... !"); */
                        SqlCommand Cmd_HDR = new SqlCommand(myquery_Light, con6_AMS);
                        Cmd_HDR.CommandTimeout = 90;
                        SqlDataReader Reader_AMS = Cmd_HDR.ExecuteReader();

                        /* MessageBox.Show(myquery_Light); */
                        /* On récupère le nombre de lignes reourné par la requête*/
                        DataTable AMAS_DT = new DataTable();
                        AMAS_DT.Load(Reader_AMS);
                        int Count_myquery_Light = AMAS_DT.Rows.Count;
                        /*   MessageBox.Show(" Le nombre de lignes retourné est : " + Count_myquery_Light.ToString()  + ""); */
                        Reader_AMS = Cmd_HDR.ExecuteReader();
                        /*  */
                        int Count = Count_myquery_Light + 1;

                        string[] Z = new string[Count];
                        string[] Y = new string[Count];
                        string[] G = new string[Count];
                        string[] F = new string[Count];
                        int b = 0;
                        /* On récupère chaque ligne dans des tableaux */
                        while (Reader_AMS.Read())
                        {
                            Z[b] = Reader_AMS.GetString(0);
                            Y[b] = Reader_AMS.GetString(1);
                            G[b] = Reader_AMS[2].ToString();
                            F[b] = Reader_AMS.GetString(3);
                            b++;
                        }
                        Reader_AMS.Close();
                        /* Insertion des valeurs selectionnées */
                        for (int j = 0; j < Count - 1; j++)
                        {
                            /*  MessageBox.Show(" Le client et ses infos sont : " + Z[j] + " , " + Y[j] + " , " + G[j] + ", " + F[j] +""); */
                            SqlConnection con_Ins_AMS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                            con_Ins_AMS.Open();
                            if (Z[j] != "" && Y[j] != "" && G[j] != "")
                            {
                                String Query_INS_AMS = " insert into [" + textBox5.Text + "AMS_I] ([CustomerID],[CurrencyCode],[SelloutAmount],[SpecialPrincingReference]) VALUES ('" + Insertion_Espaces_AMS1(Z[j]) + "', '" + Y[j] + "', '" + Insertion_Negatif(Insertion_Espaces_AMS1(Insertion_ZERO_Float(G[j].Replace(',', '.')))) + "', '" + Insertion_Espaces_AMS2(F[j]) + "' )";
                                SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_AMS, con_Ins_AMS);
                                Cmd_Remp_AMS.ExecuteNonQuery();
                                con_Ins_AMS.Close();
                            }

                        }
                    }
                    else /* Si on a des clients exclus on exéccute la requête totale */
                    {
                        SqlCommand Cmd_HDR = new SqlCommand(myquery, con6_AMS);
                        SqlDataReader Reader_AMS = Cmd_HDR.ExecuteReader();
                        /* */
                        DataTable AMAS_DT = new DataTable();
                        AMAS_DT.Load(Reader_AMS);
                        int Count_myquery_Light = AMAS_DT.Rows.Count;
                        Reader_AMS = Cmd_HDR.ExecuteReader();
                        /*  */
                        int Count = Count_myquery_Light + 1;
                        string[] Z = new string[Count];
                        string[] Y = new string[Count];
                        string[] G = new string[Count];
                        string[] F = new string[Count];
                        int b = 0;
                        while (Reader_AMS.Read())
                        {
                            Z[b] = Reader_AMS.GetString(0);
                            Y[b] = Reader_AMS.GetString(1);
                            G[b] = Reader_AMS[2].ToString();
                            F[b] = Reader_AMS.GetString(3);
                            b++;
                        }
                        Reader_AMS.Close();
                        for (int j = 0; j < Count - 1; j++)
                        {
                            SqlConnection con_Ins_AMS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                            con_Ins_AMS.Open();
                            String Query_INS_AMS = " insert into  [" + textBox5.Text + "AMS_I] ([CustomerID],[CurrencyCode],[SelloutAmount],[SpecialPrincingReference]) VALUES ('" + Insertion_Espaces_AMS1(Z[j]) + "', '" + Y[j] + "', '" + Insertion_Negatif(Insertion_Espaces_AMS1(Insertion_ZERO_Float(G[j].Replace(',', '.')))) + "', '" + Insertion_Espaces_AMS2(F[j]) + "' )";
                            SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_AMS, con_Ins_AMS);
                            Cmd_Remp_AMS.ExecuteNonQuery();
                            con_Ins_AMS.Close();
                        }

                    }
                    con6_AMS.Close();
                    richTextBox4.Text = richTextBox4.Text + "  /  " + "OK AMS_I";
                }
                return true;
            }

            catch (Exception e)
            {
                if (codeHP == CodeHPE)
                    MessageBox.Show("Remplissage AMS_E Not OK! ");
                else
                    MessageBox.Show("Remplissage AMS_I Not OK! ");
                MessageBox.Show("Exception : " + e.Message + " ");
                SqlConnection con4 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                con4.Open();
                string myquery2 = " insert into [Incidents_HP] ([Numero Etat HP] ,[Incident]) values (" + textBox5.Text.ToString() + ",'" + e.Message.ToString().Replace("'", "") + "')";
             
                SqlCommand Incident = new SqlCommand(myquery2, con4);
                Incident.ExecuteNonQuery();
                con4.Close();
                if (codeHP == CodeHPE)
                {
                    /*Notification_Mail*/
                    SendEmail(" notifhp@disway.com", "notifhp@disway.com", "Alerte HPE", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de AMS_E: " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                else
                {
                    /*Notification_Mail*/
                    SendEmail(" notifhp@disway.com", "notifhp@disway.com", "Alerte HPI", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de AMS_I: " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                return false;
            }

        }

        /* Fonction de création de la table PRS_E */
        private bool Creation_PRS_Table(string codeHP)
        {
            try
            {
                if (codeHP == CodeHPE)
                {
                    SqlConnection con5_PRS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                    con5_PRS.Open();
                    //string myquery = "CREATE TABLE [" + textBox5.Text + "PRS_E] ( SoldToCustomerID Varchar (50) , InvoicetoCustomerID  Varchar (50)  , HPProductNumber  Varchar (50) , PartnerProductNumber  Varchar (50) , DetailedSellout  Varchar (50) , SpecialPrincingreference  Varchar (50) , InvoiceNumber  Varchar (50) , InvoiceDate Varchar (50) , ShipmentDate  Varchar (50) ,InvoiceNetAmount  Varchar (50) ,CurrencyCode  Varchar (50))";
                    /**
                     *** ZHA 29/06/2015 : HP OPG
                     */

                    //string myquery = "CREATE TABLE [" + textBox5.Text + "PRS_E] ( SoldToCustomerID Varchar (50) , InvoicetoCustomerID  Varchar (50)  , HPProductNumber  Varchar (50) , PartnerProductNumber  Varchar (50) , DetailedSellout  Varchar (50) , SpecialPrincingreference  Varchar (50) , InvoiceNumber  Varchar (50) , InvoiceDate Varchar (50) , ShipmentDate  Varchar (50) ,InvoiceNetAmount  Varchar (50) ,CurrencyCode  Varchar (50),OPGId varchar(20),OPGVersion varchar(4),OPGId2 varchar(20),OPGVersion2 varchar(4))";
                    string myquery = "CREATE TABLE [" + textBox5.Text + "PRS_E] ( SoldToCustomerID Varchar (50) , InvoicetoCustomerID  Varchar (50)  , HPProductNumber  Varchar (50) , PartnerProductNumber  Varchar (50) , DetailedSellout  Varchar (50) , SpecialPrincingreference  Varchar (50) , InvoiceNumber  Varchar (50) , InvoiceDate Varchar (50) , ShipmentDate  Varchar (50) ,InvoiceNetAmount  Varchar (50) ,CurrencyCode  Varchar (50),OPGId varchar(20),OPGVersion varchar(4),OPGId2 varchar(20),OPGVersion2 varchar(4),BundleID varchar(20), ProdSerialID varchar(20), PartPurchOrderID varchar(40), HPEInvoiceNumber varchar(10), EndUserID varchar(20), ShipToCustID varchar(20), OriginCountry varchar(4), DropShipFlag varchar(4), UpFrontOPG2 varchar(20), BackendOPG2 varchar(20), UpFrontOPG3 varchar(20), BackendOPG3 varchar(20), BackendOPG4 varchar(20), DealRegID1 varchar(20), DealRegID2 varchar(20), PartInternTransID varchar(40), OriginalHPETransNumber varchar(40), PartPurshPrice varchar(20), ExtendNetCostAfterRebate varchar(20), TerritoryManager varchar(40))";

                    SqlCommand Cmd_HDR =
                        new SqlCommand(myquery, con5_PRS);
                    Cmd_HDR.ExecuteNonQuery();
                    con5_PRS.Close();
                    /*MessageBox.Show("OK PRS_E");*/
                    richTextBox3.Text = richTextBox3.Text + " / " + "OK PRS_E";
                }
                else
                {
                    SqlConnection con5_PRS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                    con5_PRS.Open();
                    //string myquery = "CREATE TABLE [" + textBox5.Text + "PRS_I] ( SoldToCustomerID Varchar (50) , InvoicetoCustomerID  Varchar (50)  , HPProductNumber  Varchar (50) , PartnerProductNumber  Varchar (50) , DetailedSellout  Varchar (50) , SpecialPrincingreference  Varchar (50) , InvoiceNumber  Varchar (50) , InvoiceDate Varchar (50) , ShipmentDate  Varchar (50) ,InvoiceNetAmount  Varchar (50) ,CurrencyCode  Varchar (50))";
                    /**
                     *** ZHA 29/06/2015 : HP OPG
                     */
                    string myquery = "CREATE TABLE [" + textBox5.Text + "PRS_I] ( SoldToCustomerID Varchar (50) , InvoicetoCustomerID  Varchar (50)  , HPProductNumber  Varchar (50) , PartnerProductNumber  Varchar (50) , DetailedSellout  Varchar (50) , SpecialPrincingreference  Varchar (50) , InvoiceNumber  Varchar (50) , InvoiceDate Varchar (50) , ShipmentDate  Varchar (50) ,InvoiceNetAmount  Varchar (50) ,CurrencyCode  Varchar (50),OPGId varchar(20),OPGVersion varchar(4),OPGId2 varchar(20),OPGVersion2 varchar(4),BundleID1 varchar(30),BundleID2 varchar(30),OPGID3 varchar(30), BundleID3 varchar(30),OPGID4 varchar(30), BundleID4 varchar(30), OPGID5 varchar(30),BundleID5 varchar(30),OPGID6 varchar(30),BundleID6 varchar(30),ShipToLocID varchar(30),HPInvoiceNo varchar(50),EndUserID varchar(30),PartPurchPrice varchar(30),PartPurchPriceCC varchar(30),PartRequestedRebateAmount varchar(30),PartnerComment varchar(120),PartnerReportedCBN varchar(30),PartnerReference varchar(50),PartnerInterTransID  varchar(50), IsDropShip  varchar(30),CustChannelPurchId  varchar(50),PurchaseAgreement varchar(30),ReporterPurchOrderID varchar(50),SuppliesTrackID varchar(60),IntercompanyFlag  varchar(30), ProdSerialIdHP  varchar(Max))";
                    SqlCommand Cmd_HDR =
                        new SqlCommand(myquery, con5_PRS);
                    Cmd_HDR.ExecuteNonQuery();
                    con5_PRS.Close();
                    /*MessageBox.Show("OK PRS_E");*/
                    richTextBox3.Text = richTextBox3.Text + " / " + "OK PRS_I";
                }
                return true;
            }
            catch( Exception e)
            {
                MessageBox.Show(e.Message);
                SqlConnection con4 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                con4.Open();
                string myquery2 = " insert into [Incidents_HP] ([Numero Etat HP] ,[Incident]) values (" + textBox5.Text.ToString() + ",'" + e.Message.ToString().Replace("'", "") + "')";
                
                SqlCommand Incident = new SqlCommand(myquery2, con4);
                Incident.ExecuteNonQuery();
                con4.Close();
                if (codeHP == CodeHPE)
                {
                    /*Notification_Mail*/
                    SendEmail("notifhp@disway.com", "notifhp@disway.com", "Alerte HPE", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de PRS_E: " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                else
                {
                    /*Notification_Mail*/
                    SendEmail("notifhp@disway.com", "notifhp@disway.com", "Alerte HPI", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de PRS_I: " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                return false;
            }
        }


        /* Fonction de remplissage de la table PRS_E */
        private bool Remplissage_PRS(string codeHP)
        {
            try
            {
                if (codeHP == CodeHPE)
                {
                    /* Selection des codes familles HPE */
                    SqlConnection Con_FHP = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
                    Con_FHP.Open();
                    String Query = "select DISTINCT [code_Famille] from Famille_HP_New where Cartes='"+codeHP+"'";
                    String Query2 = "Select count(*) from Famille_HP_New where Cartes='HPE'";
                    SqlCommand cmd_FamilleHP = new SqlCommand(Query, Con_FHP);
                    SqlCommand cmd_Count_FamilleHP = new SqlCommand(Query2, Con_FHP);
                    int Count_HP = (int)cmd_Count_FamilleHP.ExecuteScalar();
                    SqlDataReader Reader = cmd_FamilleHP.ExecuteReader();
                    string[] T = new string[Count_HP];
                    string H = "";
                    int i = 0;
                    while (Reader.Read())
                    {
                        T[i] = Reader.GetString(0);
                        i++;
                    }
                    Reader.Close();
                    for (int j = 0; j < Count_HP; j++)
                    {
                        if (j < Count_HP - 1)
                            H += "'" + T[j] + "',";
                        else
                            H += "'" + T[j] + "'";


                    }
                    Con_FHP.Close();

                    /* Selection des Clients exclus */

                    Con_FHP.Open();
                    String Query_cus = "select [Code_Client_Exclu] from Clients_Exclus";
                    String Querycus_cus_count = "Select count(*) from Clients_Exclus";
                    SqlCommand cmd_Cus_Exclus = new SqlCommand(Query_cus, Con_FHP);
                    SqlCommand cmd_Count_Clients_Exclus = new SqlCommand(Querycus_cus_count, Con_FHP);
                    int Count_Cus_Exclus = (int)cmd_Count_Clients_Exclus.ExecuteScalar();
                    SqlDataReader Reader_Cus = cmd_Cus_Exclus.ExecuteReader();
                    string[] R = new string[Count_Cus_Exclus];
                    string L = "";
                    int m = 0;
                    while (Reader_Cus.Read())
                    {
                        T[m] = Reader_Cus.GetString(0);
                        m++;
                    }
                    Reader_Cus.Close();
                    for (int n = 0; n < Count_Cus_Exclus; n++)
                    {
                        if (n < Count_Cus_Exclus - 1)
                            L += "'" + T[n] + "',";
                        else
                            L += "'" + T[n] + "'";
                    }
                    string CLT = "";
                    if (L != "") CLT = "AND (S.[N° donneur d'ordre] NOT IN (" + L + ")) ";
                    Con_FHP.Close();

                    /* Selection des Produits exclus */
                    Con_FHP.Open();
                    String Query_Prod = "select [Code_Produit_Exclu] from Produits_Exclus";
                    String Querycus_Prod_Count = "Select count(*) from Produits_Exclus";
                    SqlCommand cmdProd_Exclus = new SqlCommand(Query_Prod, Con_FHP);
                    SqlCommand cmd_Count_Prod_Exclus = new SqlCommand(Querycus_Prod_Count, Con_FHP);
                    int Count_Prod_Exclus = (int)cmd_Count_Prod_Exclus.ExecuteScalar();
                    cmdProd_Exclus.CommandTimeout = 90;
                    SqlDataReader Reader_Prod = cmdProd_Exclus.ExecuteReader();
                    string[] P = new string[Count_Prod_Exclus];
                    string M = "";
                    int s = 0;
                    while (Reader_Prod.Read())
                    {
                        P[s] = Reader_Prod.GetString(0);
                        s++;
                    }
                    Reader_Prod.Close();
                    for (int r = 0; r < Count_Prod_Exclus; r++)
                    {
                        if (r < Count_Prod_Exclus - 1)
                            M += "'" + P[r] + "',";
                        else
                            M += "'" + P[r] + "'";


                    }
                    string PRD = "";
                    if (M != "") PRD = "AND (P.[N°] NOT IN (" + M + ")) ";
                    Con_FHP.Close();

                    /* Selection des codes magazins a exclure : TUNISIE 
                    Con_FHP.Open();
                    String Query_Magaz = "select [Code Magasin] from Magasin";
                    SqlCommand cmd_Magaz = new SqlCommand(Query_Magaz, Con_FHP);
                    string Code_magazin = (string)cmd_Magaz.ExecuteScalar();
                    Con_FHP.Close(); */

                    /*Selection des magasins Exclus */
                    Con_FHP.Open();
                    String Query_Mag = "select  [Code Magasin] from Magasin";
                    String Query_Mag_Count = "Select count(*) from Magasin";
                    SqlCommand cmdMag_Exclus = new SqlCommand(Query_Mag, Con_FHP);
                    SqlCommand cmd_Count_Mag_Exclus = new SqlCommand(Query_Mag_Count, Con_FHP);
                    int Count_Mag_Exclus = (int)cmd_Count_Mag_Exclus.ExecuteScalar();
                    SqlDataReader Reader_Mag = cmdMag_Exclus.ExecuteReader();
                    string[] CX = new string[Count_Mag_Exclus];
                    string EX = "";
                    int x = 0;
                    while (Reader_Mag.Read())
                    {
                        CX[x] = Reader_Mag.GetString(0);
                        x++;
                    }
                    Reader_Mag.Close();
                    for (int r = 0; r < Count_Mag_Exclus; r++)
                    {
                        if (r < Count_Mag_Exclus - 1)
                            EX += "'" + CX[r] + "',";
                        else
                            EX += "'" + CX[r] + "'";


                    }
                    string MAG = "";
                    if (EX != "") MAG = "AND ( S.[Code Magasin] NOT IN (" + EX + ")) ";
                    Con_FHP.Close();


                    /* Fin de selection des magasins Exclus*/


                    /* Selection des informations de remplissage PRS_E */
                    SqlConnection con6_PRS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Disway;Integrated Security=True");
                    con6_PRS.Open();

                    if (L == "") /* Si la liste des clients exclus est vide on enleve la condition sur les clients exclus */
                    {
                        /*string myquery_Light =
    "SELECT S.[N° donneur d'ordre], S.[N°], CONVERT(int,ISNULL(SUM(S.Quantité),0)) AS Quantité, MAX(S.[N° lot]) AS [N° lot], S.[N° document] = CASE MAX(S.[N° lot]) WHEN '' THEN '' ELSE MAX(S.[N° document]) END, CONVERT(varchar(10),MAX(S.[Date comptabilisation]),126) AS InvoiceDate, ShipmentDate = CASE MAX(S.[N° lot]) WHEN '' THEN '' ELSE CONVERT(varchar(10),MAX(S.[Date comptabilisation]),126) END FROM dbo.___Sold_CM_BL S INNER JOIN dbo.___Product P ON S.[N°] = P.[N°] WHERE ( (P.[Code Famille] IN (" + H + "))  AND (P.[N°] NOT IN (" + M + ")) AND (S.[Date comptabilisation] BETWEEN '" + dateTimePicker1.Text + "' AND '" + dateTimePicker2.Text + "')  AND ( S.[Code Magasin] NOT IN (" + EX + "))) GROUP BY S.[N° donneur d'ordre], S.[N°] Having SUM(S.Quantité)<> 0 ORDER BY S.[N° donneur d'ordre]";
                        /**** ZHA 29/06/2015 : OPG HP
                         * */
                        //*** AURORA - Rajout de colonnes obligatoires
                          /*string myquery_Light =
    "SELECT S.[N° donneur d'ordre], S.[N°], CONVERT(int,ISNULL(SUM(S.Quantité),0)) AS Quantité, MAX(S.[N° lot]) AS [N° lot], S.[N° document], CONVERT(varchar(10),MAX(S.[Date comptabilisation]),126) AS InvoiceDate, CONVERT(varchar(10),MAX(S.[Date comptabilisation]),126) AS [Shipment Date],S.[OPG Id],S.[OPG Version],S.[Secondary OPG Id],S.[Secondary OPG Version],S.[N° BL],S.[N° ligne BL],S.[Serial No_ Tracking],S.[Promotion code],S.[Entry No_] FROM dbo.___Sold_CM_BL_OPG S INNER JOIN dbo.___Product P ON S.[N°] = P.[N°]  WHERE ( (P.[Code Famille] IN (" + H + ") " + Get_CodeSFamille(codeHP) + " AND (S.[Date comptabilisation] BETWEEN '" + dateTimePicker1.Text + "' AND '" + dateTimePicker2.Text + "')  " + PRD + MAG + ") GROUP BY S.[N° donneur d'ordre], S.[N°],S.[N° document],[OPG Id],[OPG Version],[Secondary OPG Id],[Secondary OPG Version],S.[N° BL],S.[N° ligne BL],S.[Serial No_ Tracking],S.[Promotion code],S.[Entry No_] Having SUM(S.Quantité)<> 0 ORDER BY S.[N° donneur d'ordre]"; */
                          //*** AURORA - Rajout de colonnes obligatoires

                        string myquery_Light =
   "SELECT S.[N° donneur d'ordre], S.[N°], CONVERT(int,ISNULL(SUM(S.Quantité),0)) AS Quantité, MAX(S.[N° lot]) AS [N° lot], [N° document] = CASE MAX(S.[N° lot]) WHEN '' THEN '' ELSE MAX(S.[N° document]) END, CONVERT(varchar(10),MAX(S.[Date comptabilisation]),126) AS InvoiceDate, ShipmentDate = CASE MAX(S.[N° lot]) WHEN '' THEN '' ELSE CONVERT(varchar(10),MAX(S.[Date comptabilisation]),126) END,S.[OPG Id],S.[OPG Version],S.[Secondary OPG Id],S.[Secondary OPG Version],S.[N° BL],S.[N° ligne BL],S.[Serial No_ Tracking],S.[Promotion code],S.[Entry No_] FROM dbo.___Sold_CM_BL_OPG S INNER JOIN dbo.___Product P ON S.[N°] = P.[N°] WHERE ( (P.[Code Famille] IN (" + H + ") " + Get_CodeSFamille(codeHP) + " AND (S.[Date comptabilisation] BETWEEN '" + dateTimePicker1.Text + "' AND '" + dateTimePicker2.Text + "')  " + PRD + MAG + ") GROUP BY S.[N° donneur d'ordre], S.[N°],[OPG Id],[OPG Version],[Secondary OPG Id],[Secondary OPG Version],S.[N° BL],S.[N° ligne BL],S.[Serial No_ Tracking],S.[Promotion code],S.[Entry No_] Having SUM(S.Quantité)<> 0 ORDER BY S.[N° donneur d'ordre]";

                        /*    MessageBox.Show("Exécution de la requête sans clients exclus  en cours ... !"); */
                        SqlCommand Cmd_HDR = new SqlCommand(myquery_Light, con6_PRS);
                        Cmd_HDR.CommandTimeout = 90;
                        SqlDataReader Reader_PRS = Cmd_HDR.ExecuteReader();

                        /* On récupère le nombre de lignes reourné par la requête*/
                        DataTable PRS_DT = new DataTable();
                        PRS_DT.Load(Reader_PRS);
                        int Count_myquery_Light = PRS_DT.Rows.Count;
                        /* MessageBox.Show(" Le nombre de lignes retourné est : " + Count_myquery_Light.ToString() + ""); */
                        /* Exécution correcte jusque la  */
                        Reader_PRS = Cmd_HDR.ExecuteReader();

                        int Count = Count_myquery_Light + 1;

                        string[] Z = new string[Count];
                        string[] Y = new string[Count];
                        string[] G = new string[Count];
                        string[] F = new string[Count];
                        string[] W = new string[Count];
                        string[] V = new string[Count];
                        string[] U = new string[Count];
                        /**
                         *** ZHA 29/06/2015 : HP OPG
                         */
                        string[] OPGId = new string[Count];
                        string[] OPGV = new string[Count];
                        string[] OPGId2 = new string[Count];
                        string[] OPGV2 = new string[Count];
                        string[] BL = new string[Count];
                        string[] NL_BL = new string[Count];
                        string[] codePromo = new string[Count];
                        int[] SN = new int[Count];
                        int[] EntryNo = new int[Count];

                        int b = 0;
                        /* On récupère chaque ligne dans des tableaux */
                        while (Reader_PRS.Read())
                        {
                            Z[b] = Reader_PRS.GetString(0);
                            Y[b] = Reader_PRS.GetString(1);
                            G[b] = Reader_PRS[2].ToString();
                            F[b] = Reader_PRS.GetString(3);
                            W[b] = Reader_PRS.GetString(4);
                            //MessageBox.Show("ton message");

                            V[b] = Reader_PRS[5].ToString();
                            U[b] = Reader_PRS.GetString(6);
                            /**
                             *** ZHA 29/06/2015 : HP OPG
                             */

                            OPGId[b] = Reader_PRS[7].ToString();
                            OPGV[b] = Reader_PRS[8].ToString();
                            OPGId2[b] = Reader_PRS[9].ToString();
                            OPGV2[b] = Reader_PRS[10].ToString();
                            BL[b] = Reader_PRS[11].ToString();
                            NL_BL[b] = Reader_PRS[12].ToString();
                            SN[b] = int.Parse(Reader_PRS[13].ToString());
                            codePromo[b] = Reader_PRS[14].ToString();
                            EntryNo[b] = int.Parse(Reader_PRS[15].ToString());
                            
                            b++;
                        }
                        Reader_PRS.Close();
                        con6_PRS.Close();

                        string BundleID = "";
                        string ProdSerialID = "";
                        string PartPurchOrderID = "POHP";
                        string HPEInvoiceNumber = "";
                        string EndUserID = "";
                        string ShipToCustID = "";
                        string OriginCountry = "MA";
                        string DropShipFlag = "N";
                        string UpFrontOPG2 = "";
                        string BackendOPG2 = "";
                        string UpFrontOPG3 = "";
                        string BackendOPG3 = "";
                        string BackendOPG4 = "";
                        string DealRegID1 = "";
                        string DealRegID2 = "";
                        string PartInternTransID = "";
                        string OriginalHPETransNumber = "";
                        string PartPurshPrice = "";
                        string ExtendNetCostAfterRebate = "";
                        string TerritoryManager = "";

                        /* Insertion des valeurs selectionnées */
                        for (int j = 0; j < Count - 1; j++)
                        {
                            ShipToCustID = Z[j];
                            SqlConnection con_Ins_PRS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                            
                            //String Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "')";
                            /**
                             *** ZHA 29/06/2015 : HP OPG
                             */
                            String Query_INS_PRS = "";
                            if (EntryNo[j] == 0)
                            {                               
                                //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "')";
                                //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2], [BundleID1], [BundleID2], [OPGID3], [BundleID3], [OPGID4], [BundleID4], [OPGID5], [BundleID5], [OPGID6], [BundleID6], [ShipToLocID], [HPInvoiceNo], [EndUserID], [PartPurchPrice], [PartPurchPriceCC], [PartRequestedRebateAmount], [PartnerComment], [PartnerReportedCBN], [PartnerReference], [PartnerInterTransID], [IsDropShip], [CustChannelPurchId], [PurchaseAgreement], [ReporterPurchOrderID], [SuppliesTrackID], [IntercompanyFlag], [ProdSerialIdHP]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "', '" + Insertion_Espaces_PRS_15(BundleID1) + "', '" + Insertion_Espaces_PRS_15(BundleID2) + "', '" + Insertion_Espaces_PRS_10(OPGID3) + "', '" + Insertion_Espaces_PRS_15(BundleID3) + "', '" + Insertion_Espaces_PRS_10(OPGID4) + "', '" + Insertion_Espaces_PRS_15(BundleID4) + "', '" + Insertion_Espaces_PRS_10(OPGID5) + "', '" + Insertion_Espaces_PRS_15(BundleID5) + "', '" + Insertion_Espaces_PRS_10(OPGID6) + "', '" + Insertion_Espaces_PRS_15(BundleID6) + "', '" + Insertion_Espaces_PRS_17(ShipToLocID) + "', '" + Insertion_Espaces_PRS_30(HPInvoiceNo) + "', '" + Insertion_Espaces_PRS_17(EndUserID) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3(PartPurchPriceCC) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_100(PartnerComment) + "', '" + Insertion_Espaces_PRS_15(PartnerReportedCBN) + "', '" + Insertion_Espaces_PRS_30(PartnerReference) + "', '" + Insertion_Espaces_PRS_40(PartnerInterTransID) + "', '" + Insertion_Espaces_PRS_6(IsDropShip) + "', '" + Insertion_Espaces_PRS_30(CustChannelPurchId) + "', '" + Insertion_Espaces_PRS_15(PurchaseAgreement) + "', '" + Insertion_Espaces_PRS_30(ReporterPurchOrderID) + "', '" + Insertion_Espaces_PRS_50(SuppliesTrackID) + "', '" + Insertion_Espaces_PRS_10(IntercompanyFlag) + "', '" + Insertion_Espaces_PRS_1600(ProdSerialIdHP) + "')";
                                Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2],[BundleID], [ProdSerialID], [PartPurchOrderID], [HPEInvoiceNumber], [EndUserID], [ShipToCustID], [OriginCountry], [DropShipFlag], [UpFrontOPG2], [BackendOPG2], [UpFrontOPG3], [BackendOPG3], [BackendOPG4], [DealRegID1], [DealRegID2], [PartInternTransID], [OriginalHPETransNumber], [PartPurshPrice], [ExtendNetCostAfterRebate], [TerritoryManager]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "','" + BundleID.PadRight(15, ' ') + "', '" + ProdSerialID.PadRight(15, ' ') + "', '" + PartPurchOrderID.PadRight(30, ' ') + "','" + HPEInvoiceNumber.PadRight(8, ' ') + "', '" + EndUserID.PadRight(17, ' ') + "', '" + ShipToCustID.PadRight(17, ' ') + "','" + OriginCountry + "', '" + DropShipFlag + "', '" + UpFrontOPG2.PadRight(10, ' ') + "','" + BackendOPG2.PadRight(10, ' ') + "', '" + UpFrontOPG3.PadRight(10, ' ') + "',  '" + BackendOPG3.PadRight(10, ' ') + "','" + BackendOPG4.PadRight(10, ' ') + "', '" + DealRegID1.PadRight(15, ' ') + "', '" + DealRegID2.PadRight(15, ' ') + "','" + PartInternTransID.PadRight(30, ' ') + "', '" + OriginalHPETransNumber.PadRight(30, ' ') + "', '" + PartPurshPrice.PadRight(15, ' ') + "', '" + ExtendNetCostAfterRebate.PadRight(15, ' ') + "', '" + TerritoryManager.PadRight(35, ' ') + "')";

                                con_Ins_PRS.Open();
                                SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_PRS, con_Ins_PRS);
                                Cmd_Remp_AMS.ExecuteNonQuery();
                                con_Ins_PRS.Close();
                                
                            }
                            else 
                            {
                                SqlConnection con_OPG = new SqlConnection("Data Source=WAYBI;Initial Catalog=Disway;Integrated Security=True");
                                con_OPG.Open();
                                SqlConnection con_OPG2 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Disway;Integrated Security=True");
                                con_OPG2.Open();

                                /*********** OPG 1 ************/
                                string R_OPG1 = "select [OPG Id],[OPG Version] from __OPG_HP_PRS_Primary where ILE_No=" + EntryNo[j] + " and [Item No_]='" + Y[j] + "'";
                                SqlCommand Cmd_OPG1 = new SqlCommand(R_OPG1, con_OPG);
                                Cmd_OPG1.CommandTimeout = 90;
                                SqlDataReader Reader_OPG1 = Cmd_OPG1.ExecuteReader();
                                /* On récupère le nombre de lignes reourné par la requête*/
                                bool OPG1 = Reader_OPG1.HasRows;

                                /*********** OPG 2 ************/
                                string R_OPG2 = "select [OPG Id 2],[OPG Version 2] from __OPG_HP_PRS_SecondeV2 where ILE_No=" + EntryNo[j] + " and [Item No_]='" + Y[j] + "'";
                                SqlCommand Cmd_OPG2 = new SqlCommand(R_OPG2, con_OPG2);
                                Cmd_OPG2.CommandTimeout = 90;
                                SqlDataReader Reader_OPG2 = Cmd_OPG2.ExecuteReader();
                                /* On récupère le nombre de lignes reourné par la requête*/
                                bool OPG2 = Reader_OPG2.HasRows;

                                if (!OPG1 && !OPG2)
                                {
                                    //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "')";
                                    //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2], [BundleID1], [BundleID2], [OPGID3], [BundleID3], [OPGID4], [BundleID4], [OPGID5], [BundleID5], [OPGID6], [BundleID6], [ShipToLocID], [HPInvoiceNo], [EndUserID], [PartPurchPrice], [PartPurchPriceCC], [PartRequestedRebateAmount], [PartnerComment], [PartnerReportedCBN], [PartnerReference], [PartnerInterTransID], [IsDropShip], [CustChannelPurchId], [PurchaseAgreement], [ReporterPurchOrderID], [SuppliesTrackID], [IntercompanyFlag], [ProdSerialIdHP]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "', '" + Insertion_Espaces_PRS_15(BundleID1) + "', '" + Insertion_Espaces_PRS_15(BundleID2) + "', '" + Insertion_Espaces_PRS_10(OPGID3) + "', '" + Insertion_Espaces_PRS_15(BundleID3) + "', '" + Insertion_Espaces_PRS_10(OPGID4) + "', '" + Insertion_Espaces_PRS_15(BundleID4) + "', '" + Insertion_Espaces_PRS_10(OPGID5) + "', '" + Insertion_Espaces_PRS_15(BundleID5) + "', '" + Insertion_Espaces_PRS_10(OPGID6) + "', '" + Insertion_Espaces_PRS_15(BundleID6) + "', '" + Insertion_Espaces_PRS_17(ShipToLocID) + "', '" + Insertion_Espaces_PRS_30(HPInvoiceNo) + "', '" + Insertion_Espaces_PRS_17(EndUserID) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3(PartPurchPriceCC) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_100(PartnerComment) + "', '" + Insertion_Espaces_PRS_15(PartnerReportedCBN) + "', '" + Insertion_Espaces_PRS_30(PartnerReference) + "', '" + Insertion_Espaces_PRS_40(PartnerInterTransID) + "', '" + Insertion_Espaces_PRS_6(IsDropShip) + "', '" + Insertion_Espaces_PRS_30(CustChannelPurchId) + "', '" + Insertion_Espaces_PRS_15(PurchaseAgreement) + "', '" + Insertion_Espaces_PRS_30(ReporterPurchOrderID) + "', '" + Insertion_Espaces_PRS_50(SuppliesTrackID) + "', '" + Insertion_Espaces_PRS_10(IntercompanyFlag) + "', '" + Insertion_Espaces_PRS_1600(ProdSerialIdHP) + "')";
                                    Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2],[BundleID], [ProdSerialID], [PartPurchOrderID], [HPEInvoiceNumber], [EndUserID], [ShipToCustID], [OriginCountry], [DropShipFlag], [UpFrontOPG2], [BackendOPG2], [UpFrontOPG3], [BackendOPG3], [BackendOPG4], [DealRegID1], [DealRegID2], [PartInternTransID], [OriginalHPETransNumber], [PartPurshPrice], [ExtendNetCostAfterRebate], [TerritoryManager]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "','" + BundleID.PadRight(15, ' ') + "', '" + ProdSerialID.PadRight(15, ' ') + "', '" + PartPurchOrderID.PadRight(30, ' ') + "','" + HPEInvoiceNumber.PadRight(8, ' ') + "', '" + EndUserID.PadRight(17, ' ') + "', '" + ShipToCustID.PadRight(17, ' ') + "','" + OriginCountry + "', '" + DropShipFlag + "', '" + UpFrontOPG2.PadRight(10, ' ') + "','" + BackendOPG2.PadRight(10, ' ') + "', '" + UpFrontOPG3.PadRight(10, ' ') + "',  '" + BackendOPG3.PadRight(10, ' ') + "','" + BackendOPG4.PadRight(10, ' ') + "', '" + DealRegID1.PadRight(15, ' ') + "', '" + DealRegID2.PadRight(15, ' ') + "','" + PartInternTransID.PadRight(30, ' ') + "', '" + OriginalHPETransNumber.PadRight(30, ' ') + "', '" + PartPurshPrice.PadRight(15, ' ') + "', '" + ExtendNetCostAfterRebate.PadRight(15, ' ') + "', '" + TerritoryManager.PadRight(35, ' ') + "')";

                                    con_Ins_PRS.Open();
                                    SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_PRS, con_Ins_PRS);
                                    Cmd_Remp_AMS.ExecuteNonQuery();
                                    con_Ins_PRS.Close();
                                }
                                if (OPG1 && OPG2)
                                {
                                    while (Reader_OPG1.Read() && Reader_OPG2.Read())
                                    {
                                        if (Reader_OPG1[0].ToString() == "")
                                        {
                                            OPGId[j] = Reader_OPG2[0].ToString();
                                            OPGV[j] = Reader_OPG2[1].ToString();
                                            OPGId2[j] = Reader_OPG1[0].ToString();
                                            OPGV2[j] = Reader_OPG1[1].ToString();
                                        }
                                        else
                                        {
                                            OPGId[j] = Reader_OPG1[0].ToString();
                                            OPGV[j] = Reader_OPG1[1].ToString();
                                            OPGId2[j] = Reader_OPG2[0].ToString();
                                            OPGV2[j] = Reader_OPG2[1].ToString();
                                        }
                                        //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "')";
                                        //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2], [BundleID1], [BundleID2], [OPGID3], [BundleID3], [OPGID4], [BundleID4], [OPGID5], [BundleID5], [OPGID6], [BundleID6], [ShipToLocID], [HPInvoiceNo], [EndUserID], [PartPurchPrice], [PartPurchPriceCC], [PartRequestedRebateAmount], [PartnerComment], [PartnerReportedCBN], [PartnerReference], [PartnerInterTransID], [IsDropShip], [CustChannelPurchId], [PurchaseAgreement], [ReporterPurchOrderID], [SuppliesTrackID], [IntercompanyFlag], [ProdSerialIdHP]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "', '" + Insertion_Espaces_PRS_15(BundleID1) + "', '" + Insertion_Espaces_PRS_15(BundleID2) + "', '" + Insertion_Espaces_PRS_10(OPGID3) + "', '" + Insertion_Espaces_PRS_15(BundleID3) + "', '" + Insertion_Espaces_PRS_10(OPGID4) + "', '" + Insertion_Espaces_PRS_15(BundleID4) + "', '" + Insertion_Espaces_PRS_10(OPGID5) + "', '" + Insertion_Espaces_PRS_15(BundleID5) + "', '" + Insertion_Espaces_PRS_10(OPGID6) + "', '" + Insertion_Espaces_PRS_15(BundleID6) + "', '" + Insertion_Espaces_PRS_17(ShipToLocID) + "', '" + Insertion_Espaces_PRS_30(HPInvoiceNo) + "', '" + Insertion_Espaces_PRS_17(EndUserID) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3(PartPurchPriceCC) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_100(PartnerComment) + "', '" + Insertion_Espaces_PRS_15(PartnerReportedCBN) + "', '" + Insertion_Espaces_PRS_30(PartnerReference) + "', '" + Insertion_Espaces_PRS_40(PartnerInterTransID) + "', '" + Insertion_Espaces_PRS_6(IsDropShip) + "', '" + Insertion_Espaces_PRS_30(CustChannelPurchId) + "', '" + Insertion_Espaces_PRS_15(PurchaseAgreement) + "', '" + Insertion_Espaces_PRS_30(ReporterPurchOrderID) + "', '" + Insertion_Espaces_PRS_50(SuppliesTrackID) + "', '" + Insertion_Espaces_PRS_10(IntercompanyFlag) + "', '" + Insertion_Espaces_PRS_1600(ProdSerialIdHP) + "')";
                                        Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2],[BundleID], [ProdSerialID], [PartPurchOrderID], [HPEInvoiceNumber], [EndUserID], [ShipToCustID], [OriginCountry], [DropShipFlag], [UpFrontOPG2], [BackendOPG2], [UpFrontOPG3], [BackendOPG3], [BackendOPG4], [DealRegID1], [DealRegID2], [PartInternTransID], [OriginalHPETransNumber], [PartPurshPrice], [ExtendNetCostAfterRebate], [TerritoryManager]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "','" + BundleID.PadRight(15, ' ') + "', '" + ProdSerialID.PadRight(15, ' ') + "', '" + PartPurchOrderID.PadRight(30, ' ') + "','" + HPEInvoiceNumber.PadRight(8, ' ') + "', '" + EndUserID.PadRight(17, ' ') + "', '" + ShipToCustID.PadRight(17, ' ') + "','" + OriginCountry + "', '" + DropShipFlag + "', '" + UpFrontOPG2.PadRight(10, ' ') + "','" + BackendOPG2.PadRight(10, ' ') + "', '" + UpFrontOPG3.PadRight(10, ' ') + "',  '" + BackendOPG3.PadRight(10, ' ') + "','" + BackendOPG4.PadRight(10, ' ') + "', '" + DealRegID1.PadRight(15, ' ') + "', '" + DealRegID2.PadRight(15, ' ') + "','" + PartInternTransID.PadRight(30, ' ') + "', '" + OriginalHPETransNumber.PadRight(30, ' ') + "', '" + PartPurshPrice.PadRight(15, ' ') + "', '" + ExtendNetCostAfterRebate.PadRight(15, ' ') + "', '" + TerritoryManager.PadRight(35, ' ') + "')";

                                        con_Ins_PRS.Open();
                                        SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_PRS, con_Ins_PRS);
                                        Cmd_Remp_AMS.ExecuteNonQuery();
                                        con_Ins_PRS.Close();
                                    }     
                                }
                                if (OPG1 && !OPG2)
                                {
                                    while (Reader_OPG1.Read())
                                    {
                                        //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(Reader_OPG1[0].ToString()) + "','" + Insertion_Espaces_PRS_3(Reader_OPG1[1].ToString()) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "')";
                                        //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2], [BundleID1], [BundleID2], [OPGID3], [BundleID3], [OPGID4], [BundleID4], [OPGID5], [BundleID5], [OPGID6], [BundleID6], [ShipToLocID], [HPInvoiceNo], [EndUserID], [PartPurchPrice], [PartPurchPriceCC], [PartRequestedRebateAmount], [PartnerComment], [PartnerReportedCBN], [PartnerReference], [PartnerInterTransID], [IsDropShip], [CustChannelPurchId], [PurchaseAgreement], [ReporterPurchOrderID], [SuppliesTrackID], [IntercompanyFlag], [ProdSerialIdHP]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(Reader_OPG1[0].ToString()) + "','" + Insertion_Espaces_PRS_3(Reader_OPG1[1].ToString()) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "', '" + Insertion_Espaces_PRS_15(BundleID1) + "', '" + Insertion_Espaces_PRS_15(BundleID2) + "', '" + Insertion_Espaces_PRS_10(OPGID3) + "', '" + Insertion_Espaces_PRS_15(BundleID3) + "', '" + Insertion_Espaces_PRS_10(OPGID4) + "', '" + Insertion_Espaces_PRS_15(BundleID4) + "', '" + Insertion_Espaces_PRS_10(OPGID5) + "', '" + Insertion_Espaces_PRS_15(BundleID5) + "', '" + Insertion_Espaces_PRS_10(OPGID6) + "', '" + Insertion_Espaces_PRS_15(BundleID6) + "', '" + Insertion_Espaces_PRS_17(ShipToLocID) + "', '" + Insertion_Espaces_PRS_30(HPInvoiceNo) + "', '" + Insertion_Espaces_PRS_17(EndUserID) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3(PartPurchPriceCC) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_100(PartnerComment) + "', '" + Insertion_Espaces_PRS_15(PartnerReportedCBN) + "', '" + Insertion_Espaces_PRS_30(PartnerReference) + "', '" + Insertion_Espaces_PRS_40(PartnerInterTransID) + "', '" + Insertion_Espaces_PRS_6(IsDropShip) + "', '" + Insertion_Espaces_PRS_30(CustChannelPurchId) + "', '" + Insertion_Espaces_PRS_15(PurchaseAgreement) + "', '" + Insertion_Espaces_PRS_30(ReporterPurchOrderID) + "', '" + Insertion_Espaces_PRS_50(SuppliesTrackID) + "', '" + Insertion_Espaces_PRS_10(IntercompanyFlag) + "', '" + Insertion_Espaces_PRS_1600(ProdSerialIdHP) + "')";
                                        Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2],[BundleID], [ProdSerialID], [PartPurchOrderID], [HPEInvoiceNumber], [EndUserID], [ShipToCustID], [OriginCountry], [DropShipFlag], [UpFrontOPG2], [BackendOPG2], [UpFrontOPG3], [BackendOPG3], [BackendOPG4], [DealRegID1], [DealRegID2], [PartInternTransID], [OriginalHPETransNumber], [PartPurshPrice], [ExtendNetCostAfterRebate], [TerritoryManager]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(Reader_OPG1[0].ToString()) + "','" + Insertion_Espaces_PRS_3(Reader_OPG1[1].ToString()) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "','" + BundleID.PadRight(15, ' ') + "', '" + ProdSerialID.PadRight(15, ' ') + "', '" + PartPurchOrderID.PadRight(30, ' ') + "','" + HPEInvoiceNumber.PadRight(8, ' ') + "', '" + EndUserID.PadRight(17, ' ') + "', '" + ShipToCustID.PadRight(17, ' ') + "','" + OriginCountry + "', '" + DropShipFlag + "', '" + UpFrontOPG2.PadRight(10, ' ') + "','" + BackendOPG2.PadRight(10, ' ') + "', '" + UpFrontOPG3.PadRight(10, ' ') + "',  '" + BackendOPG3.PadRight(10, ' ') + "','" + BackendOPG4.PadRight(10, ' ') + "', '" + DealRegID1.PadRight(15, ' ') + "', '" + DealRegID2.PadRight(15, ' ') + "','" + PartInternTransID.PadRight(30, ' ') + "', '" + OriginalHPETransNumber.PadRight(30, ' ') + "', '" + PartPurshPrice.PadRight(15, ' ') + "', '" + ExtendNetCostAfterRebate.PadRight(15, ' ') + "', '" + TerritoryManager.PadRight(35, ' ') + "')";

                                        con_Ins_PRS.Open();
                                        SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_PRS, con_Ins_PRS);
                                        Cmd_Remp_AMS.ExecuteNonQuery();
                                        con_Ins_PRS.Close();
                                    }
                                }
                                if (!OPG1 && OPG2)
                                {
                                    while (Reader_OPG2.Read())
                                    {
                                        //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(Reader_OPG2[0].ToString()) + "','" + Insertion_Espaces_PRS_3(Reader_OPG2[1].ToString()) + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "')";
                                        //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2], [BundleID1], [BundleID2], [OPGID3], [BundleID3], [OPGID4], [BundleID4], [OPGID5], [BundleID5], [OPGID6], [BundleID6], [ShipToLocID], [HPInvoiceNo], [EndUserID], [PartPurchPrice], [PartPurchPriceCC], [PartRequestedRebateAmount], [PartnerComment], [PartnerReportedCBN], [PartnerReference], [PartnerInterTransID], [IsDropShip], [CustChannelPurchId], [PurchaseAgreement], [ReporterPurchOrderID], [SuppliesTrackID], [IntercompanyFlag], [ProdSerialIdHP]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(Reader_OPG2[0].ToString()) + "','" + Insertion_Espaces_PRS_3(Reader_OPG2[1].ToString()) + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "')";
                                        Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2],[BundleID], [ProdSerialID], [PartPurchOrderID], [HPEInvoiceNumber], [EndUserID], [ShipToCustID], [OriginCountry], [DropShipFlag], [UpFrontOPG2], [BackendOPG2], [UpFrontOPG3], [BackendOPG3], [BackendOPG4], [DealRegID1], [DealRegID2], [PartInternTransID], [OriginalHPETransNumber], [PartPurshPrice], [ExtendNetCostAfterRebate], [TerritoryManager]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(Reader_OPG2[0].ToString()) + "','" + Insertion_Espaces_PRS_3(Reader_OPG2[1].ToString()) + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + BundleID.PadRight(15, ' ') + "', '" + ProdSerialID.PadRight(15, ' ') + "', '" + PartPurchOrderID.PadRight(30, ' ') + "','" + HPEInvoiceNumber.PadRight(8, ' ') + "', '" + EndUserID.PadRight(17, ' ') + "', '" + ShipToCustID.PadRight(17, ' ') + "','" + OriginCountry + "', '" + DropShipFlag + "', '" + UpFrontOPG2.PadRight(10, ' ') + "','" + BackendOPG2.PadRight(10, ' ') + "', '" + UpFrontOPG3.PadRight(10, ' ') + "',  '" + BackendOPG3.PadRight(10, ' ') + "','" + BackendOPG4.PadRight(10, ' ') + "', '" + DealRegID1.PadRight(15, ' ') + "', '" + DealRegID2.PadRight(15, ' ') + "','" + PartInternTransID.PadRight(30, ' ') + "', '" + OriginalHPETransNumber.PadRight(30, ' ') + "', '" + PartPurshPrice.PadRight(15, ' ') + "', '" + ExtendNetCostAfterRebate.PadRight(15, ' ') + "', '" + TerritoryManager.PadRight(35, ' ') + "')";

                                        con_Ins_PRS.Open();
                                        SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_PRS, con_Ins_PRS);
                                        Cmd_Remp_AMS.ExecuteNonQuery();
                                        con_Ins_PRS.Close();
                                    }
                                }
                                Reader_OPG1.Close();
                                Reader_OPG2.Close();
                                con_OPG.Close();
                                con_OPG2.Close();
                            }
                            /* Remplissage SNT */
                            if (SN[j] == 1 && codePromo[j] != "")
                            {
                                Remplissage_SNT(codeHP, BL[j], NL_BL[j]);
                            }
                            
                        }
                    }
                    /* SI la liste des clients exclus n'est pas vide on exécute la requête total */
                    else
                    {
                        /*string myquery =
    "SELECT S.[N° donneur d'ordre], S.[N°], CONVERT(int,ISNULL(SUM(S.Quantité),0)) AS Quantité, MAX(S.[N° lot]) AS [N° lot], S.[N° document] = CASE MAX(S.[N° lot]) WHEN '' THEN '' ELSE MAX(S.[N° document]) END, MAX(S.[Date comptabilisation]) AS InvoiceDate, ShipmentDate = CASE MAX(S.[N° lot]) WHEN '' THEN '' ELSE CONVERT(char, MAX(S.[Date comptabilisation])) END FROM dbo.___Sold_CM_BL S INNER JOIN dbo.___Product P ON S.[N°] = P.[N°] WHERE (((P.[Code Famille] IN (" + H + ")) AND (S.[N° donneur d'ordre] NOT IN (" + L + ")) AND (P.[N°] NOT IN (" + M + ")) AND (S.[Date comptabilisation] BETWEEN '" + dateTimePicker1.Text + "' AND '" + dateTimePicker2.Text + "') AND (S.[Code Magasin] NOT IN (" + EX + ") )))  GROUP BY S.[N° donneur d'ordre], S.[N°] Having SUM(S.Quantité)<> 0 ORDER BY S.[N° donneur d'ordre]";
                        /**** ZHA 29/06/2015 : OPG HP
                         */ 
                        /*string myquery =
    "SELECT S.[N° donneur d'ordre], S.[N°], CONVERT(int,ISNULL(SUM(S.Quantité),0)) AS Quantité, MAX(S.[N° lot]) AS [N° lot], S.[N° document], CONVERT(varchar(10),MAX(S.[Date comptabilisation]),126) AS InvoiceDate, ShipmentDate = CASE MAX(S.[N° lot]) WHEN '' THEN '' ELSE CONVERT(varchar(10),MAX(S.[Date comptabilisation]),126) END,[OPG Id],[OPG Version],[Secondary OPG Id],[Secondary OPG Version],S.[N° BL],S.[N° ligne BL],S.[Serial No_ Tracking],S.[Promotion code],S.[Entry No_] FROM dbo.___Sold_CM_BL_OPG S INNER JOIN dbo.___Product P ON S.[N°] = P.[N°] WHERE (P.[Code Famille] IN (" + H + ") " + Get_CodeSFamille(codeHP) + " AND (S.[Date comptabilisation] BETWEEN '" + dateTimePicker1.Text + "' AND '" + dateTimePicker2.Text + "') " + CLT + PRD + MAG + " GROUP BY S.[N° donneur d'ordre], S.[N°],[OPG Id],[OPG Version],[Secondary OPG Id],[Secondary OPG Version],S.[N° BL],S.[N° ligne BL],S.[Serial No_ Tracking],S.[Promotion code],S.[Entry No_] Having SUM(S.Quantité)<> 0 ORDER BY S.[N° donneur d'ordre]"; */

                        string myquery =
    "SELECT S.[N° donneur d'ordre], S.[N°], CONVERT(int,ISNULL(SUM(S.Quantité),0)) AS Quantité, MAX(S.[N° lot]) AS [N° lot], [N° document] = CASE MAX(S.[N° lot]) WHEN '' THEN '' ELSE MAX(S.[N° document]) END, CONVERT(varchar(10),MAX(S.[Date comptabilisation]),126) AS InvoiceDate, ShipmentDate = CASE MAX(S.[N° lot]) WHEN '' THEN '' ELSE CONVERT(varchar(10),MAX(S.[Date comptabilisation]),126) END,[OPG Id],[OPG Version],[Secondary OPG Id],[Secondary OPG Version],S.[N° BL],S.[N° ligne BL],S.[Serial No_ Tracking],S.[Promotion code],S.[Entry No_] FROM dbo.___Sold_CM_BL_OPG S INNER JOIN dbo.___Product P ON S.[N°] = P.[N°] WHERE (P.[Code Famille] IN (" + H + ") " + Get_CodeSFamille(codeHP) + " AND (S.[Date comptabilisation] BETWEEN '" + dateTimePicker1.Text + "' AND '" + dateTimePicker2.Text + "') " + CLT + PRD + MAG + " GROUP BY S.[N° donneur d'ordre], S.[N°],[OPG Id],[OPG Version],[Secondary OPG Id],[Secondary OPG Version],S.[N° BL],S.[N° ligne BL],S.[Serial No_ Tracking],S.[Promotion code],S.[Entry No_] Having SUM(S.Quantité)<> 0 ORDER BY S.[N° donneur d'ordre]";
                        
                        SqlCommand Cmd_HDR = new SqlCommand(myquery, con6_PRS);
                        Cmd_HDR.CommandTimeout = 90;
                        SqlDataReader Reader_PRS = Cmd_HDR.ExecuteReader();   /**************** bloque à ce niveau ********/
                        /* */
                        DataTable PRS_DT = new DataTable();
                        PRS_DT.Load(Reader_PRS);
                        int Count_myquery = PRS_DT.Rows.Count;
                        Reader_PRS = Cmd_HDR.ExecuteReader();
                        /*  */
                        int Count = Count_myquery + 1;
                        string[] Z = new string[Count];
                        string[] Y = new string[Count];
                        string[] G = new string[Count];
                        string[] F = new string[Count];
                        string[] W = new string[Count];
                        string[] V = new string[Count];
                        string[] U = new string[Count];
                        /**
                         *** ZHA 29/06/2015 : HP OPG
                         */
                        string[] OPGId = new string[Count];
                        string[] OPGV = new string[Count];
                        string[] OPGId2 = new string[Count];
                        string[] OPGV2 = new string[Count];
                        string[] BL = new string[Count];
                        string[] NL_BL = new string[Count];
                        string[] codePromo = new string[Count];
                        int[] SN = new int[Count];
                        int[] EntryNo = new int[Count];
                        int b = 0;
                        while (Reader_PRS.Read())
                        {
                            Z[b] = Reader_PRS.GetString(0);
                            Y[b] = Reader_PRS.GetString(1);
                            G[b] = Reader_PRS[2].ToString();
                            F[b] = Reader_PRS.GetString(3);
                            W[b] = Reader_PRS.GetString(4);
                            V[b] = Reader_PRS[5].ToString();
                            U[b] = Reader_PRS.GetString(6);
                            /**
                            *** ZHA 29/06/2015 : HP OPG
                            */
                            OPGId[b] = Reader_PRS[7].ToString(); //.GetString(7);
                            OPGV[b] = Reader_PRS[8].ToString();
                            OPGId2[b] = Reader_PRS[9].ToString();
                            OPGV2[b] = Reader_PRS[10].ToString();
                            BL[b] = Reader_PRS[11].ToString();
                            NL_BL[b] = Reader_PRS[12].ToString();
                            SN[b] = int.Parse(Reader_PRS[13].ToString());
                            codePromo[b] = Reader_PRS[14].ToString();
                            EntryNo[b] = int.Parse(Reader_PRS[15].ToString());
                            b++;
                        }
                        Reader_PRS.Close();

                        string BundleID = "";
                        string ProdSerialID = "";
                        string PartPurchOrderID = "POHP";
                        string HPEInvoiceNumber = "";
                        string EndUserID = "";
                        string ShipToCustID = "";
                        string OriginCountry = "MA";
                        string DropShipFlag = "N";
                        string UpFrontOPG2 = "";
                        string BackendOPG2 = "";
                        string UpFrontOPG3 = "";
                        string BackendOPG3 = "";
                        string BackendOPG4 = "";
                        string DealRegID1 = "";
                        string DealRegID2 = "";
                        string PartInternTransID = "";
                        string OriginalHPETransNumber = "";
                        string PartPurshPrice = "";
                        string ExtendNetCostAfterRebate = "";
                        string TerritoryManager = "";

                        for (int j = 0; j < Count - 1; j++)
                        {
                            ShipToCustID = Z[j];
                            SqlConnection con_Ins_PRS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                            //con_Ins_PRS.Open();
                            //String Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "')";
                            /**
                             *** ZHA 29/06/2015 : HP OPG
                             */
                            String Query_INS_PRS = "";
                            if (EntryNo[j] == 0)
                            {
                                //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "')";
                                //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2], [BundleID1], [BundleID2], [OPGID3], [BundleID3], [OPGID4], [BundleID4], [OPGID5], [BundleID5], [OPGID6], [BundleID6], [ShipToLocID], [HPInvoiceNo], [EndUserID], [PartPurchPrice], [PartPurchPriceCC], [PartRequestedRebateAmount], [PartnerComment], [PartnerReportedCBN], [PartnerReference], [PartnerInterTransID], [IsDropShip], [CustChannelPurchId], [PurchaseAgreement], [ReporterPurchOrderID], [SuppliesTrackID], [IntercompanyFlag], [ProdSerialIdHP]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "', '" + Insertion_Espaces_PRS_15(BundleID1) + "', '" + Insertion_Espaces_PRS_15(BundleID2) + "', '" + Insertion_Espaces_PRS_10(OPGID3) + "', '" + Insertion_Espaces_PRS_15(BundleID3) + "', '" + Insertion_Espaces_PRS_10(OPGID4) + "', '" + Insertion_Espaces_PRS_15(BundleID4) + "', '" + Insertion_Espaces_PRS_10(OPGID5) + "', '" + Insertion_Espaces_PRS_15(BundleID5) + "', '" + Insertion_Espaces_PRS_10(OPGID6) + "', '" + Insertion_Espaces_PRS_15(BundleID6) + "', '" + Insertion_Espaces_PRS_17(ShipToLocID) + "', '" + Insertion_Espaces_PRS_30(HPInvoiceNo) + "', '" + Insertion_Espaces_PRS_17(EndUserID) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3(PartPurchPriceCC) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_100(PartnerComment) + "', '" + Insertion_Espaces_PRS_15(PartnerReportedCBN) + "', '" + Insertion_Espaces_PRS_30(PartnerReference) + "', '" + Insertion_Espaces_PRS_40(PartnerInterTransID) + "', '" + Insertion_Espaces_PRS_6(IsDropShip) + "', '" + Insertion_Espaces_PRS_30(CustChannelPurchId) + "', '" + Insertion_Espaces_PRS_15(PurchaseAgreement) + "', '" + Insertion_Espaces_PRS_30(ReporterPurchOrderID) + "', '" + Insertion_Espaces_PRS_50(SuppliesTrackID) + "', '" + Insertion_Espaces_PRS_10(IntercompanyFlag) + "', '" + Insertion_Espaces_PRS_1600(ProdSerialIdHP) + "')";
                                Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2],[BundleID], [ProdSerialID], [PartPurchOrderID], [HPEInvoiceNumber], [EndUserID], [ShipToCustID], [OriginCountry], [DropShipFlag], [UpFrontOPG2], [BackendOPG2], [UpFrontOPG3], [BackendOPG3], [BackendOPG4], [DealRegID1], [DealRegID2], [PartInternTransID], [OriginalHPETransNumber], [PartPurshPrice], [ExtendNetCostAfterRebate], [TerritoryManager]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "','" + BundleID.PadRight(15, ' ') + "', '" + ProdSerialID.PadRight(15, ' ') + "', '" + PartPurchOrderID.PadRight(30, ' ') + "','" + HPEInvoiceNumber.PadRight(8, ' ') + "', '" + EndUserID.PadRight(17, ' ') + "', '" + ShipToCustID.PadRight(17, ' ') + "','" + OriginCountry + "', '" + DropShipFlag + "', '" + UpFrontOPG2.PadRight(10, ' ') + "','" + BackendOPG2.PadRight(10, ' ') + "', '" + UpFrontOPG3.PadRight(10, ' ') + "',  '" + BackendOPG3.PadRight(10, ' ') + "','" + BackendOPG4.PadRight(10, ' ') + "', '" + DealRegID1.PadRight(15, ' ') + "', '" + DealRegID2.PadRight(15, ' ') + "','" + PartInternTransID.PadRight(30, ' ') + "', '" + OriginalHPETransNumber.PadRight(30, ' ') + "', '" + PartPurshPrice.PadRight(15, ' ') + "', '" + ExtendNetCostAfterRebate.PadRight(15, ' ') + "', '" + TerritoryManager.PadRight(35, ' ') + "')";

                                con_Ins_PRS.Open();
                                SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_PRS, con_Ins_PRS);
                                Cmd_Remp_AMS.ExecuteNonQuery();
                                con_Ins_PRS.Close();

                            }
                            else
                            {
                                SqlConnection con_OPG = new SqlConnection("Data Source=WAYBI;Initial Catalog=Disway;Integrated Security=True");
                                con_OPG.Open();
                                SqlConnection con_OPG2 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Disway;Integrated Security=True");
                                con_OPG2.Open();
                                /*********** OPG 1 ************/
                                string R_OPG1 = "select [OPG Id],[OPG Version] from __OPG_HP_PRS_Primary where ILE_No=" + EntryNo[j] + " and [Item No_]='" + Y[j] + "' group by [OPG Id],[OPG Version]";
                                SqlCommand Cmd_OPG1 = new SqlCommand(R_OPG1, con_OPG);
                                Cmd_OPG1.CommandTimeout = 90;
                                SqlDataReader Reader_OPG1 = Cmd_OPG1.ExecuteReader();
                                bool OPG1 = Reader_OPG1.HasRows;


                                /*********** OPG 2 ************/
                                string R_OPG2 = "select [OPG Id 2],[OPG Version 2] from __OPG_HP_PRS_SecondeV2 where ILE_No=" + EntryNo[j] + " and [Item No_]='" + Y[j] + "'";
                                SqlCommand Cmd_OPG2 = new SqlCommand(R_OPG2, con_OPG2);
                                Cmd_OPG2.CommandTimeout = 90;
                                SqlDataReader Reader_OPG2 = Cmd_OPG2.ExecuteReader();
                                bool OPG2 = Reader_OPG2.HasRows;


                                if (!OPG1 && !OPG2)
                                {
                                    //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "')";
                                    //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2], [BundleID1], [BundleID2], [OPGID3], [BundleID3], [OPGID4], [BundleID4], [OPGID5], [BundleID5], [OPGID6], [BundleID6], [ShipToLocID], [HPInvoiceNo], [EndUserID], [PartPurchPrice], [PartPurchPriceCC], [PartRequestedRebateAmount], [PartnerComment], [PartnerReportedCBN], [PartnerReference], [PartnerInterTransID], [IsDropShip], [CustChannelPurchId], [PurchaseAgreement], [ReporterPurchOrderID], [SuppliesTrackID], [IntercompanyFlag], [ProdSerialIdHP]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "', '" + Insertion_Espaces_PRS_15(BundleID1) + "', '" + Insertion_Espaces_PRS_15(BundleID2) + "', '" + Insertion_Espaces_PRS_10(OPGID3) + "', '" + Insertion_Espaces_PRS_15(BundleID3) + "', '" + Insertion_Espaces_PRS_10(OPGID4) + "', '" + Insertion_Espaces_PRS_15(BundleID4) + "', '" + Insertion_Espaces_PRS_10(OPGID5) + "', '" + Insertion_Espaces_PRS_15(BundleID5) + "', '" + Insertion_Espaces_PRS_10(OPGID6) + "', '" + Insertion_Espaces_PRS_15(BundleID6) + "', '" + Insertion_Espaces_PRS_17(ShipToLocID) + "', '" + Insertion_Espaces_PRS_30(HPInvoiceNo) + "', '" + Insertion_Espaces_PRS_17(EndUserID) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3(PartPurchPriceCC) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_100(PartnerComment) + "', '" + Insertion_Espaces_PRS_15(PartnerReportedCBN) + "', '" + Insertion_Espaces_PRS_30(PartnerReference) + "', '" + Insertion_Espaces_PRS_40(PartnerInterTransID) + "', '" + Insertion_Espaces_PRS_6(IsDropShip) + "', '" + Insertion_Espaces_PRS_30(CustChannelPurchId) + "', '" + Insertion_Espaces_PRS_15(PurchaseAgreement) + "', '" + Insertion_Espaces_PRS_30(ReporterPurchOrderID) + "', '" + Insertion_Espaces_PRS_50(SuppliesTrackID) + "', '" + Insertion_Espaces_PRS_10(IntercompanyFlag) + "', '" + Insertion_Espaces_PRS_1600(ProdSerialIdHP) + "')";
                                    Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2],[BundleID], [ProdSerialID], [PartPurchOrderID], [HPEInvoiceNumber], [EndUserID], [ShipToCustID], [OriginCountry], [DropShipFlag], [UpFrontOPG2], [BackendOPG2], [UpFrontOPG3], [BackendOPG3], [BackendOPG4], [DealRegID1], [DealRegID2], [PartInternTransID], [OriginalHPETransNumber], [PartPurshPrice], [ExtendNetCostAfterRebate], [TerritoryManager]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "','" + BundleID.PadRight(15, ' ') + "', '" + ProdSerialID.PadRight(15, ' ') + "', '" + PartPurchOrderID.PadRight(30, ' ') + "','" + HPEInvoiceNumber.PadRight(8, ' ') + "', '" + EndUserID.PadRight(17, ' ') + "', '" + ShipToCustID.PadRight(17, ' ') + "','" + OriginCountry + "', '" + DropShipFlag + "', '" + UpFrontOPG2.PadRight(10, ' ') + "','" + BackendOPG2.PadRight(10, ' ') + "', '" + UpFrontOPG3.PadRight(10, ' ') + "',  '" + BackendOPG3.PadRight(10, ' ') + "','" + BackendOPG4.PadRight(10, ' ') + "', '" + DealRegID1.PadRight(15, ' ') + "', '" + DealRegID2.PadRight(15, ' ') + "','" + PartInternTransID.PadRight(30, ' ') + "', '" + OriginalHPETransNumber.PadRight(30, ' ') + "', '" + PartPurshPrice.PadRight(15, ' ') + "', '" + ExtendNetCostAfterRebate.PadRight(15, ' ') + "', '" + TerritoryManager.PadRight(35, ' ') + "')";
                                    
                                    con_Ins_PRS.Open();
                                    SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_PRS, con_Ins_PRS);
                                    Cmd_Remp_AMS.ExecuteNonQuery();
                                    con_Ins_PRS.Close();
                                }
                                if (OPG1 && OPG2)
                                {

                                    while (Reader_OPG1.Read() && Reader_OPG2.Read())
                                    {
                                        if (Reader_OPG1[0].ToString() == "")
                                        {
                                            OPGId[j] = Reader_OPG2[0].ToString();
                                            OPGV[j] = Reader_OPG2[1].ToString();
                                            OPGId2[j] = Reader_OPG1[0].ToString();
                                            OPGV2[j] = Reader_OPG1[1].ToString();
                                        }
                                        else
                                        {
                                            OPGId[j] = Reader_OPG1[0].ToString();
                                            OPGV[j] = Reader_OPG1[1].ToString();
                                            OPGId2[j] = Reader_OPG2[0].ToString();
                                            OPGV2[j] = Reader_OPG2[1].ToString();
                                        }
                                        //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "')";
                                        //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2], [BundleID1], [BundleID2], [OPGID3], [BundleID3], [OPGID4], [BundleID4], [OPGID5], [BundleID5], [OPGID6], [BundleID6], [ShipToLocID], [HPInvoiceNo], [EndUserID], [PartPurchPrice], [PartPurchPriceCC], [PartRequestedRebateAmount], [PartnerComment], [PartnerReportedCBN], [PartnerReference], [PartnerInterTransID], [IsDropShip], [CustChannelPurchId], [PurchaseAgreement], [ReporterPurchOrderID], [SuppliesTrackID], [IntercompanyFlag], [ProdSerialIdHP]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "', '" + Insertion_Espaces_PRS_15(BundleID1) + "', '" + Insertion_Espaces_PRS_15(BundleID2) + "', '" + Insertion_Espaces_PRS_10(OPGID3) + "', '" + Insertion_Espaces_PRS_15(BundleID3) + "', '" + Insertion_Espaces_PRS_10(OPGID4) + "', '" + Insertion_Espaces_PRS_15(BundleID4) + "', '" + Insertion_Espaces_PRS_10(OPGID5) + "', '" + Insertion_Espaces_PRS_15(BundleID5) + "', '" + Insertion_Espaces_PRS_10(OPGID6) + "', '" + Insertion_Espaces_PRS_15(BundleID6) + "', '" + Insertion_Espaces_PRS_17(ShipToLocID) + "', '" + Insertion_Espaces_PRS_30(HPInvoiceNo) + "', '" + Insertion_Espaces_PRS_17(EndUserID) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3(PartPurchPriceCC) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_100(PartnerComment) + "', '" + Insertion_Espaces_PRS_15(PartnerReportedCBN) + "', '" + Insertion_Espaces_PRS_30(PartnerReference) + "', '" + Insertion_Espaces_PRS_40(PartnerInterTransID) + "', '" + Insertion_Espaces_PRS_6(IsDropShip) + "', '" + Insertion_Espaces_PRS_30(CustChannelPurchId) + "', '" + Insertion_Espaces_PRS_15(PurchaseAgreement) + "', '" + Insertion_Espaces_PRS_30(ReporterPurchOrderID) + "', '" + Insertion_Espaces_PRS_50(SuppliesTrackID) + "', '" + Insertion_Espaces_PRS_10(IntercompanyFlag) + "', '" + Insertion_Espaces_PRS_1600(ProdSerialIdHP) + "')";
                                        Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2],[BundleID], [ProdSerialID], [PartPurchOrderID], [HPEInvoiceNumber], [EndUserID], [ShipToCustID], [OriginCountry], [DropShipFlag], [UpFrontOPG2], [BackendOPG2], [UpFrontOPG3], [BackendOPG3], [BackendOPG4], [DealRegID1], [DealRegID2], [PartInternTransID], [OriginalHPETransNumber], [PartPurshPrice], [ExtendNetCostAfterRebate], [TerritoryManager]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "','" + BundleID.PadRight(15, ' ') + "', '" + ProdSerialID.PadRight(15, ' ') + "', '" + PartPurchOrderID.PadRight(30, ' ') + "','" + HPEInvoiceNumber.PadRight(8, ' ') + "', '" + EndUserID.PadRight(17, ' ') + "', '" + ShipToCustID.PadRight(17, ' ') + "','" + OriginCountry + "', '" + DropShipFlag + "', '" + UpFrontOPG2.PadRight(10, ' ') + "','" + BackendOPG2.PadRight(10, ' ') + "', '" + UpFrontOPG3.PadRight(10, ' ') + "',  '" + BackendOPG3.PadRight(10, ' ') + "','" + BackendOPG4.PadRight(10, ' ') + "', '" + DealRegID1.PadRight(15, ' ') + "', '" + DealRegID2.PadRight(15, ' ') + "','" + PartInternTransID.PadRight(30, ' ') + "', '" + OriginalHPETransNumber.PadRight(30, ' ') + "', '" + PartPurshPrice.PadRight(15, ' ') + "', '" + ExtendNetCostAfterRebate.PadRight(15, ' ') + "', '" + TerritoryManager.PadRight(35, ' ') + "')";
                                        
                                        con_Ins_PRS.Open();
                                        SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_PRS, con_Ins_PRS);
                                        Cmd_Remp_AMS.ExecuteNonQuery();
                                        con_Ins_PRS.Close();
                                    }

                                }
                                if (OPG1 && !OPG2)
                                {
                                    while (Reader_OPG1.Read())
                                    {
                                        //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(Reader_OPG1[0].ToString()) + "','" + Insertion_Espaces_PRS_3(Reader_OPG1[1].ToString()) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "')";
                                        //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2], [BundleID1], [BundleID2], [OPGID3], [BundleID3], [OPGID4], [BundleID4], [OPGID5], [BundleID5], [OPGID6], [BundleID6], [ShipToLocID], [HPInvoiceNo], [EndUserID], [PartPurchPrice], [PartPurchPriceCC], [PartRequestedRebateAmount], [PartnerComment], [PartnerReportedCBN], [PartnerReference], [PartnerInterTransID], [IsDropShip], [CustChannelPurchId], [PurchaseAgreement], [ReporterPurchOrderID], [SuppliesTrackID], [IntercompanyFlag], [ProdSerialIdHP]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(Reader_OPG1[0].ToString()) + "','" + Insertion_Espaces_PRS_3(Reader_OPG1[1].ToString()) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "', '" + Insertion_Espaces_PRS_15(BundleID1) + "', '" + Insertion_Espaces_PRS_15(BundleID2) + "', '" + Insertion_Espaces_PRS_10(OPGID3) + "', '" + Insertion_Espaces_PRS_15(BundleID3) + "', '" + Insertion_Espaces_PRS_10(OPGID4) + "', '" + Insertion_Espaces_PRS_15(BundleID4) + "', '" + Insertion_Espaces_PRS_10(OPGID5) + "', '" + Insertion_Espaces_PRS_15(BundleID5) + "', '" + Insertion_Espaces_PRS_10(OPGID6) + "', '" + Insertion_Espaces_PRS_15(BundleID6) + "', '" + Insertion_Espaces_PRS_17(ShipToLocID) + "', '" + Insertion_Espaces_PRS_30(HPInvoiceNo) + "', '" + Insertion_Espaces_PRS_17(EndUserID) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3(PartPurchPriceCC) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_100(PartnerComment) + "', '" + Insertion_Espaces_PRS_15(PartnerReportedCBN) + "', '" + Insertion_Espaces_PRS_30(PartnerReference) + "', '" + Insertion_Espaces_PRS_40(PartnerInterTransID) + "', '" + Insertion_Espaces_PRS_6(IsDropShip) + "', '" + Insertion_Espaces_PRS_30(CustChannelPurchId) + "', '" + Insertion_Espaces_PRS_15(PurchaseAgreement) + "', '" + Insertion_Espaces_PRS_30(ReporterPurchOrderID) + "', '" + Insertion_Espaces_PRS_50(SuppliesTrackID) + "', '" + Insertion_Espaces_PRS_10(IntercompanyFlag) + "', '" + Insertion_Espaces_PRS_1600(ProdSerialIdHP) + "')";
                                        Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2],[BundleID], [ProdSerialID], [PartPurchOrderID], [HPEInvoiceNumber], [EndUserID], [ShipToCustID], [OriginCountry], [DropShipFlag], [UpFrontOPG2], [BackendOPG2], [UpFrontOPG3], [BackendOPG3], [BackendOPG4], [DealRegID1], [DealRegID2], [PartInternTransID], [OriginalHPETransNumber], [PartPurshPrice], [ExtendNetCostAfterRebate], [TerritoryManager]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(Reader_OPG1[0].ToString()) + "','" + Insertion_Espaces_PRS_3(Reader_OPG1[1].ToString()) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "','" + BundleID.PadRight(15, ' ') + "', '" + ProdSerialID.PadRight(15, ' ') + "', '" + PartPurchOrderID.PadRight(30, ' ') + "','" + HPEInvoiceNumber.PadRight(8, ' ') + "', '" + EndUserID.PadRight(17, ' ') + "', '" + ShipToCustID.PadRight(17, ' ') + "','" + OriginCountry + "', '" + DropShipFlag + "', '" + UpFrontOPG2.PadRight(10, ' ') + "','" + BackendOPG2.PadRight(10, ' ') + "', '" + UpFrontOPG3.PadRight(10, ' ') + "',  '" + BackendOPG3.PadRight(10, ' ') + "','" + BackendOPG4.PadRight(10, ' ') + "', '" + DealRegID1.PadRight(15, ' ') + "', '" + DealRegID2.PadRight(15, ' ') + "','" + PartInternTransID.PadRight(30, ' ') + "', '" + OriginalHPETransNumber.PadRight(30, ' ') + "', '" + PartPurshPrice.PadRight(15, ' ') + "', '" + ExtendNetCostAfterRebate.PadRight(15, ' ') + "', '" + TerritoryManager.PadRight(35, ' ') + "')";
                                        
                                        con_Ins_PRS.Open();
                                        SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_PRS, con_Ins_PRS);
                                        Cmd_Remp_AMS.ExecuteNonQuery();
                                        con_Ins_PRS.Close();
                                    }
                                }
                                if (!OPG1 && OPG2)
                                {
                                    while (Reader_OPG2.Read())
                                    {
                                        //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(Reader_OPG2[0].ToString()) + "','" + Insertion_Espaces_PRS_3(Reader_OPG2[1].ToString()) + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "')";
                                        //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2], [BundleID1], [BundleID2], [OPGID3], [BundleID3], [OPGID4], [BundleID4], [OPGID5], [BundleID5], [OPGID6], [BundleID6], [ShipToLocID], [HPInvoiceNo], [EndUserID], [PartPurchPrice], [PartPurchPriceCC], [PartRequestedRebateAmount], [PartnerComment], [PartnerReportedCBN], [PartnerReference], [PartnerInterTransID], [IsDropShip], [CustChannelPurchId], [PurchaseAgreement], [ReporterPurchOrderID], [SuppliesTrackID], [IntercompanyFlag], [ProdSerialIdHP]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(Reader_OPG2[0].ToString()) + "','" + Insertion_Espaces_PRS_3(Reader_OPG2[1].ToString()) + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "', '" + Insertion_Espaces_PRS_15(BundleID1) + "', '" + Insertion_Espaces_PRS_15(BundleID2) + "', '" + Insertion_Espaces_PRS_10(OPGID3) + "', '" + Insertion_Espaces_PRS_15(BundleID3) + "', '" + Insertion_Espaces_PRS_10(OPGID4) + "', '" + Insertion_Espaces_PRS_15(BundleID4) + "', '" + Insertion_Espaces_PRS_10(OPGID5) + "', '" + Insertion_Espaces_PRS_15(BundleID5) + "', '" + Insertion_Espaces_PRS_10(OPGID6) + "', '" + Insertion_Espaces_PRS_15(BundleID6) + "', '" + Insertion_Espaces_PRS_17(ShipToLocID) + "', '" + Insertion_Espaces_PRS_30(HPInvoiceNo) + "', '" + Insertion_Espaces_PRS_17(EndUserID) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3(PartPurchPriceCC) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_100(PartnerComment) + "', '" + Insertion_Espaces_PRS_15(PartnerReportedCBN) + "', '" + Insertion_Espaces_PRS_30(PartnerReference) + "', '" + Insertion_Espaces_PRS_40(PartnerInterTransID) + "', '" + Insertion_Espaces_PRS_6(IsDropShip) + "', '" + Insertion_Espaces_PRS_30(CustChannelPurchId) + "', '" + Insertion_Espaces_PRS_15(PurchaseAgreement) + "', '" + Insertion_Espaces_PRS_30(ReporterPurchOrderID) + "', '" + Insertion_Espaces_PRS_50(SuppliesTrackID) + "', '" + Insertion_Espaces_PRS_10(IntercompanyFlag) + "', '" + Insertion_Espaces_PRS_1600(ProdSerialIdHP) + "')";
                                        Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_E] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2],[BundleID], [ProdSerialID], [PartPurchOrderID], [HPEInvoiceNumber], [EndUserID], [ShipToCustID], [OriginCountry], [DropShipFlag], [UpFrontOPG2], [BackendOPG2], [UpFrontOPG3], [BackendOPG3], [BackendOPG4], [DealRegID1], [DealRegID2], [PartInternTransID], [OriginalHPETransNumber], [PartPurshPrice], [ExtendNetCostAfterRebate], [TerritoryManager]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(Reader_OPG2[0].ToString()) + "','" + Insertion_Espaces_PRS_3(Reader_OPG2[1].ToString()) + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + BundleID.PadRight(15, ' ') + "', '" + ProdSerialID.PadRight(15, ' ') + "', '" + PartPurchOrderID.PadRight(30, ' ') + "','" + HPEInvoiceNumber.PadRight(8, ' ') + "', '" + EndUserID.PadRight(17, ' ') + "', '" + ShipToCustID.PadRight(17, ' ') + "','" + OriginCountry + "', '" + DropShipFlag + "', '" + UpFrontOPG2.PadRight(10, ' ') + "','" + BackendOPG2.PadRight(10, ' ') + "', '" + UpFrontOPG3.PadRight(10, ' ') + "',  '" + BackendOPG3.PadRight(10, ' ') + "','" + BackendOPG4.PadRight(10, ' ') + "', '" + DealRegID1.PadRight(15, ' ') + "', '" + DealRegID2.PadRight(15, ' ') + "','" + PartInternTransID.PadRight(30, ' ') + "', '" + OriginalHPETransNumber.PadRight(30, ' ') + "', '" + PartPurshPrice.PadRight(15, ' ') + "', '" + ExtendNetCostAfterRebate.PadRight(15, ' ') + "', '" + TerritoryManager.PadRight(35, ' ') + "')";
                                        
                                        con_Ins_PRS.Open();
                                        SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_PRS, con_Ins_PRS);
                                        Cmd_Remp_AMS.ExecuteNonQuery();
                                        con_Ins_PRS.Close();
                                    }
                                }
                                Reader_OPG1.Close();
                                Reader_OPG2.Close();
                                con_OPG.Close();
                                con_OPG2.Close();
                            }
                            /* Remplissage SNT */
                            if (SN[j] == 1 && codePromo[j] != "")
                            {
                                Remplissage_SNT(codeHP, BL[j], NL_BL[j]);
                            }
                        }

                    }
                    con6_PRS.Close();
                    richTextBox4.Text = richTextBox4.Text + " / " + "OK PRS_E";
                }
                else
                {
                    /******************** Selection des codes familles HPI ******************************/
                    SqlConnection Con_FHP = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
                    Con_FHP.Open();
                    String Query = "select DISTINCT [code_Famille] from Famille_HP_New where Cartes='" + codeHP +"'";
                    String Query2 = "Select count(*) from Famille_HP_New where Cartes='"+codeHP+"'";
                    SqlCommand cmd_FamilleHP = new SqlCommand(Query, Con_FHP);
                    SqlCommand cmd_Count_FamilleHP = new SqlCommand(Query2, Con_FHP);
                    int Count_HP = (int)cmd_Count_FamilleHP.ExecuteScalar();
                    SqlDataReader Reader = cmd_FamilleHP.ExecuteReader();
                    string[] T = new string[Count_HP];
                    string H = "";
                    int i = 0;
                    while (Reader.Read())
                    {
                        T[i] = Reader.GetString(0);
                        i++;
                    }
                    Reader.Close();
                    for (int j = 0; j < Count_HP; j++)
                    {
                        if (j < Count_HP - 1)
                            H += "'" + T[j] + "',";
                        else
                            H += "'" + T[j] + "'";


                    }
                    Con_FHP.Close();

                    /* Selection des Clients exclus */

                    Con_FHP.Open();
                    String Query_cus = "select [Code_Client_Exclu] from Clients_Exclus";
                    String Querycus_cus_count = "Select count(*) from Clients_Exclus";
                    SqlCommand cmd_Cus_Exclus = new SqlCommand(Query_cus, Con_FHP);
                    SqlCommand cmd_Count_Clients_Exclus = new SqlCommand(Querycus_cus_count, Con_FHP);
                    int Count_Cus_Exclus = (int)cmd_Count_Clients_Exclus.ExecuteScalar();
                    SqlDataReader Reader_Cus = cmd_Cus_Exclus.ExecuteReader();
                    string[] R = new string[Count_Cus_Exclus];
                    string L = "";
                    int m = 0;
                    while (Reader_Cus.Read())
                    {
                        T[m] = Reader_Cus.GetString(0);
                        m++;
                    }
                    Reader_Cus.Close();
                    for (int n = 0; n < Count_Cus_Exclus; n++)
                    {
                        if (n < Count_Cus_Exclus - 1)
                            L += "'" + T[n] + "',";
                        else
                            L += "'" + T[n] + "'";
                    }
                    string CLT = "";
                    if (L != "") CLT = "AND (S.[N° donneur d'ordre] NOT IN (" + L + ")) ";
                    Con_FHP.Close();

                    /* Selection des Produits exclus */
                    Con_FHP.Open();
                    String Query_Prod = "select  [Code_Produit_Exclu] from Produits_Exclus";
                    String Querycus_Prod_Count = "Select count(*) from Produits_Exclus";
                    SqlCommand cmdProd_Exclus = new SqlCommand(Query_Prod, Con_FHP);
                    SqlCommand cmd_Count_Prod_Exclus = new SqlCommand(Querycus_Prod_Count, Con_FHP);
                    int Count_Prod_Exclus = (int)cmd_Count_Prod_Exclus.ExecuteScalar();
                    cmdProd_Exclus.CommandTimeout = 90;
                    SqlDataReader Reader_Prod = cmdProd_Exclus.ExecuteReader();
                    string[] P = new string[Count_Prod_Exclus];
                    string M = "";
                    int s = 0;
                    while (Reader_Prod.Read())
                    {
                        P[s] = Reader_Prod.GetString(0);
                        s++;
                    }
                    Reader_Prod.Close();
                    for (int r = 0; r < Count_Prod_Exclus; r++)
                    {
                        if (r < Count_Prod_Exclus - 1)
                            M += "'" + P[r] + "',";
                        else
                            M += "'" + P[r] + "'";


                    }
                    string PRD = "";
                    if (M != "") PRD = "AND (P.[N°] NOT IN (" + M + ")) ";
                    Con_FHP.Close();

                    /* Selection des codes magazins a exclure : TUNISIE 
                    Con_FHP.Open();
                    String Query_Magaz = "select [Code Magasin] from Magasin";
                    SqlCommand cmd_Magaz = new SqlCommand(Query_Magaz, Con_FHP);
                    string Code_magazin = (string)cmd_Magaz.ExecuteScalar();
                    Con_FHP.Close(); */

                    /*Selection des magasins Exclus */
                    Con_FHP.Open();
                    String Query_Mag = "select  [Code Magasin] from Magasin";
                    String Query_Mag_Count = "Select count(*) from Magasin";
                    SqlCommand cmdMag_Exclus = new SqlCommand(Query_Mag, Con_FHP);
                    SqlCommand cmd_Count_Mag_Exclus = new SqlCommand(Query_Mag_Count, Con_FHP);
                    int Count_Mag_Exclus = (int)cmd_Count_Mag_Exclus.ExecuteScalar();
                    SqlDataReader Reader_Mag = cmdMag_Exclus.ExecuteReader();
                    string[] CX = new string[Count_Mag_Exclus];
                    string EX = "";
                    int x = 0;
                    while (Reader_Mag.Read())
                    {
                        CX[x] = Reader_Mag.GetString(0);
                        x++;
                    }
                    Reader_Mag.Close();
                    for (int r = 0; r < Count_Mag_Exclus; r++)
                    {
                        if (r < Count_Mag_Exclus - 1)
                            EX += "'" + CX[r] + "',";
                        else
                            EX += "'" + CX[r] + "'";


                    }
                    string MAG = "";
                    if (EX != "") MAG = "AND ( S.[Code Magasin] NOT IN (" + EX + ")) ";
                    Con_FHP.Close();


                    /* Fin de selection des magasins Exclus*/


                    /* Selection des informations de remplissage PRS_I */
                    SqlConnection con6_PRS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Disway;Integrated Security=True");
                    con6_PRS.Open();

                    if (L == "") /* Si la liste des clients exclus est vide on enleve la condition sur les clients exclus */
                    {
                        /*string myquery_Light =
    "SELECT S.[N° donneur d'ordre], S.[N°], CONVERT(int,ISNULL(SUM(S.Quantité),0)) AS Quantité, MAX(S.[N° lot]) AS [N° lot], S.[N° document] = CASE MAX(S.[N° lot]) WHEN '' THEN '' ELSE MAX(S.[N° document]) END, CONVERT(varchar(10),MAX(S.[Date comptabilisation]),126) AS InvoiceDate, ShipmentDate = CASE MAX(S.[N° lot]) WHEN '' THEN '' ELSE CONVERT(varchar(10),MAX(S.[Date comptabilisation]),126) END FROM dbo.___Sold_CM_BL S INNER JOIN dbo.___Product P ON S.[N°] = P.[N°] WHERE ( (P.[Code Famille] IN (" + H + ") )  AND (P.[N°] NOT IN (" + M + ")) AND (S.[Date comptabilisation] BETWEEN '" + dateTimePicker1.Text + "' AND '" + dateTimePicker2.Text + "')  AND ( S.[Code Magasin] NOT IN (" + EX + "))) GROUP BY S.[N° donneur d'ordre], S.[N°] Having SUM(S.Quantité)<> 0 ORDER BY S.[N° donneur d'ordre]";
                        /**** ZHA 29/06/2015 : OPG HP */
                          string myquery_Light =
    "SELECT S.[N° donneur d'ordre], S.[N°], CONVERT(int,ISNULL(SUM(S.Quantité),0)) AS Quantité, MAX(S.[N° lot]) AS [N° lot], S.[N° document], CONVERT(varchar(10),MAX(S.[Date comptabilisation]),126) AS InvoiceDate, CONVERT(varchar(10),MAX(S.[Date comptabilisation]),126) AS ShipmentDate ,S.[OPG Id],S.[OPG Version],S.[Secondary OPG Id],S.[Secondary OPG Version],S.[N° BL],S.[N° ligne BL],S.[Serial No_ Tracking],S.[Promotion code],S.[Entry No_] FROM dbo.___Sold_CM_BL_OPG S INNER JOIN dbo.___Product P ON S.[N°] = P.[N°] WHERE ( (P.[Code Famille] IN (" + H + ")) " + Get_CodeSFamille(codeHP) + " AND (S.[Date comptabilisation] BETWEEN '" + dateTimePicker1.Text + "' AND '" + dateTimePicker2.Text + "')  " + PRD + MAG + ") GROUP BY S.[N° donneur d'ordre],S.[N° document], S.[N°],[OPG Id],[OPG Version],[Secondary OPG Id],[Secondary OPG Version],S.[N° BL],S.[N° ligne BL],S.[Serial No_ Tracking],S.[Promotion code],S.[Entry No_] Having SUM(S.Quantité)<> 0 ORDER BY S.[N° donneur d'ordre]";
                         
                        /*    MessageBox.Show("Exécution de la requête sans clients exclus  en cours ... !"); */
                        SqlCommand Cmd_HDR = new SqlCommand(myquery_Light, con6_PRS);
                        Cmd_HDR.CommandTimeout = 90;
                        SqlDataReader Reader_PRS = Cmd_HDR.ExecuteReader();

                        /* On récupère le nombre de lignes reourné par la requête*/
                        DataTable PRS_DT = new DataTable();
                        PRS_DT.Load(Reader_PRS);
                        int Count_myquery_Light = PRS_DT.Rows.Count;
                        /* MessageBox.Show(" Le nombre de lignes retourné est : " + Count_myquery_Light.ToString() + ""); */
                        /* Exécution correcte jusque la  */
                        Reader_PRS = Cmd_HDR.ExecuteReader();

                        int Count = Count_myquery_Light + 1;

                        string[] Z = new string[Count];
                        string[] Y = new string[Count];
                        string[] G = new string[Count];
                        string[] F = new string[Count];
                        string[] W = new string[Count];
                        string[] V = new string[Count];
                        string[] U = new string[Count];
                        /**
                         *** ZHA 29/06/2015 : HP OPG
                         */
                        string[] OPGId = new string[Count];
                        string[] OPGV = new string[Count];
                        string[] OPGId2 = new string[Count];
                        string[] OPGV2 = new string[Count];
                        string[] BL = new string[Count];
                        string[] NL_BL = new string[Count];
                        string[] codePromo = new string[Count];
                        int[] SN = new int[Count];
                        int[] EntryNo = new int[Count];
                        int b = 0;
                        while (Reader_PRS.Read())
                        {
                            Z[b] = Reader_PRS.GetString(0);
                            Y[b] = Reader_PRS.GetString(1);
                            G[b] = Reader_PRS[2].ToString();
                            F[b] = Reader_PRS.GetString(3);
                            W[b] = Reader_PRS.GetString(4);
                            V[b] = Reader_PRS[5].ToString();
                            U[b] = Reader_PRS.GetString(6);
                            /**
                            *** ZHA 29/06/2015 : HP OPG
                            */
                            OPGId[b] = Reader_PRS[7].ToString(); //.GetString(7);
                            OPGV[b] = Reader_PRS[8].ToString();
                            OPGId2[b] = Reader_PRS[9].ToString();
                            OPGV2[b] = Reader_PRS[10].ToString();
                            BL[b] = Reader_PRS[11].ToString();
                            NL_BL[b] = Reader_PRS[12].ToString();
                            SN[b] = int.Parse(Reader_PRS[13].ToString());
                            codePromo[b] = Reader_PRS[14].ToString();
                            EntryNo[b] = int.Parse(Reader_PRS[15].ToString());
                            b++;
                        }
                        Reader_PRS.Close();
                        con6_PRS.Close();

                        string BundleID1 = " ";
                        string BundleID2 = " ";
                        string OPGID3 = " ";
                        string BundleID3 = " ";
                        string OPGID4 = " ";
                        string BundleID4 = " ";
                        string OPGID5 = " ";
                        string BundleID5 = " ";
                        string OPGID6 = " ";
                        string BundleID6 = " ";
                        
                        string HPInvoiceNo = " ";
                        string EndUserID = " ";
                        //string PartPurchPrice = " ";
                        string PartPurchPriceCC = " ";
                        //string PartRequestedRebateAmount = " ";
                        string PartnerComment = " ";
                        string PartnerReportedCBN = " ";
                        string PartnerReference = " ";
                        string PartnerInterTransID = " ";
                        string IsDropShip = " ";
                        string CustChannelPurchId = " ";
                        string PurchaseAgreement = " ";
                        string ReporterPurchOrderID = " ";
                        string SuppliesTrackID = " ";
                        string IntercompanyFlag = " ";
                        string ProdSerialIdHP = " ";
                            
                        /* Insertion des valeurs selectionnées */
                        for (int j = 0; j < Count - 1; j++)
                        {
                            string ShipToLocID = Z[j];
                            SqlConnection con_Ins_PRS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                            //con_Ins_PRS.Open();
                            //String Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_I] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "')";
                            /**
                             *** ZHA 29/06/2015 : HP OPG
                             */
                            String Query_INS_PRS = "";
                            if (EntryNo[j] == 0)
                            {
                                //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_I] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "')";
                                Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_I] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2], [BundleID1], [BundleID2], [OPGID3], [BundleID3], [OPGID4], [BundleID4], [OPGID5], [BundleID5], [OPGID6], [BundleID6], [ShipToLocID], [HPInvoiceNo], [EndUserID], [PartPurchPrice], [PartPurchPriceCC], [PartRequestedRebateAmount], [PartnerComment], [PartnerReportedCBN], [PartnerReference], [PartnerInterTransID], [IsDropShip], [CustChannelPurchId], [PurchaseAgreement], [ReporterPurchOrderID], [SuppliesTrackID], [IntercompanyFlag], [ProdSerialIdHP]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "', '" + Insertion_Espaces_PRS_15(BundleID1) + "', '" + Insertion_Espaces_PRS_15(BundleID2) + "', '" + Insertion_Espaces_PRS_10(OPGID3) + "', '" + Insertion_Espaces_PRS_15(BundleID3) + "', '" + Insertion_Espaces_PRS_10(OPGID4) + "', '" + Insertion_Espaces_PRS_15(BundleID4) + "', '" + Insertion_Espaces_PRS_10(OPGID5) + "', '" + Insertion_Espaces_PRS_15(BundleID5) + "', '" + Insertion_Espaces_PRS_10(OPGID6) + "', '" + Insertion_Espaces_PRS_15(BundleID6) + "', '" + Insertion_Espaces_PRS_17(ShipToLocID) + "', '" + Insertion_Espaces_PRS_30(HPInvoiceNo) + "', '" + Insertion_Espaces_PRS_17(EndUserID) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3(PartPurchPriceCC) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_100(PartnerComment) + "', '" + Insertion_Espaces_PRS_15(PartnerReportedCBN) + "', '" + Insertion_Espaces_PRS_30(PartnerReference) + "', '" + Insertion_Espaces_PRS_40(PartnerInterTransID) + "', '" + Insertion_Espaces_PRS_6(IsDropShip) + "', '" + Insertion_Espaces_PRS_30(CustChannelPurchId) + "', '" + Insertion_Espaces_PRS_15(PurchaseAgreement) + "', '" + Insertion_Espaces_PRS_30(ReporterPurchOrderID) + "', '" + Insertion_Espaces_PRS_50(SuppliesTrackID) + "', '" + Insertion_Espaces_PRS_10(IntercompanyFlag) + "', '" + Insertion_Espaces_PRS_1600(ProdSerialIdHP) + "')";
                                con_Ins_PRS.Open();
                                SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_PRS, con_Ins_PRS);
                                Cmd_Remp_AMS.ExecuteNonQuery();
                                con_Ins_PRS.Close();

                            }
                            else
                            {
                                SqlConnection con_OPG = new SqlConnection("Data Source=WAYBI;Initial Catalog=Disway;Integrated Security=True");
                                con_OPG.Open();
                                SqlConnection con_OPG2 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Disway;Integrated Security=True");
                                con_OPG2.Open();
                                /*********** OPG 1 ************/
                                string R_OPG1 = "select [OPG Id],[OPG Version] from __OPG_HP_PRS_Primary where ILE_No=" + EntryNo[j] + " and [Item No_]='" + Y[j] + "'";
                                SqlCommand Cmd_OPG1 = new SqlCommand(R_OPG1, con_OPG);
                                Cmd_OPG1.CommandTimeout = 90;
                                SqlDataReader Reader_OPG1 = Cmd_OPG1.ExecuteReader();
                                /* On récupère le nombre de lignes reourné par la requête*/
                                bool OPG1 = Reader_OPG1.HasRows;

                                /*********** OPG 2 ************/
                                string R_OPG2 = "select [OPG Id 2],[OPG Version 2] from __OPG_HP_PRS_SecondeV2 where ILE_No=" + EntryNo[j] + " and [Item No_]='" + Y[j] + "'";
                                SqlCommand Cmd_OPG2 = new SqlCommand(R_OPG2, con_OPG2);
                                Cmd_OPG2.CommandTimeout = 90;
                                SqlDataReader Reader_OPG2 = Cmd_OPG2.ExecuteReader();
                                /* On récupère le nombre de lignes reourné par la requête*/
                                bool OPG2 = Reader_OPG2.HasRows;

                                if (!OPG1 && !OPG2)
                                {
                                    //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_I] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "')";
                                    Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_I] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2], [BundleID1], [BundleID2], [OPGID3], [BundleID3], [OPGID4], [BundleID4], [OPGID5], [BundleID5], [OPGID6], [BundleID6], [ShipToLocID], [HPInvoiceNo], [EndUserID], [PartPurchPrice], [PartPurchPriceCC], [PartRequestedRebateAmount], [PartnerComment], [PartnerReportedCBN], [PartnerReference], [PartnerInterTransID], [IsDropShip], [CustChannelPurchId], [PurchaseAgreement], [ReporterPurchOrderID], [SuppliesTrackID], [IntercompanyFlag], [ProdSerialIdHP]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "', '" + Insertion_Espaces_PRS_15(BundleID1) + "', '" + Insertion_Espaces_PRS_15(BundleID2) + "', '" + Insertion_Espaces_PRS_10(OPGID3) + "', '" + Insertion_Espaces_PRS_15(BundleID3) + "', '" + Insertion_Espaces_PRS_10(OPGID4) + "', '" + Insertion_Espaces_PRS_15(BundleID4) + "', '" + Insertion_Espaces_PRS_10(OPGID5) + "', '" + Insertion_Espaces_PRS_15(BundleID5) + "', '" + Insertion_Espaces_PRS_10(OPGID6) + "', '" + Insertion_Espaces_PRS_15(BundleID6) + "', '" + Insertion_Espaces_PRS_17(ShipToLocID) + "', '" + Insertion_Espaces_PRS_30(HPInvoiceNo) + "', '" + Insertion_Espaces_PRS_17(EndUserID) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3(PartPurchPriceCC) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_100(PartnerComment) + "', '" + Insertion_Espaces_PRS_15(PartnerReportedCBN) + "', '" + Insertion_Espaces_PRS_30(PartnerReference) + "', '" + Insertion_Espaces_PRS_40(PartnerInterTransID) + "', '" + Insertion_Espaces_PRS_6(IsDropShip) + "', '" + Insertion_Espaces_PRS_30(CustChannelPurchId) + "', '" + Insertion_Espaces_PRS_15(PurchaseAgreement) + "', '" + Insertion_Espaces_PRS_30(ReporterPurchOrderID) + "', '" + Insertion_Espaces_PRS_50(SuppliesTrackID) + "', '" + Insertion_Espaces_PRS_10(IntercompanyFlag) + "', '" + Insertion_Espaces_PRS_1600(ProdSerialIdHP) + "')";
                                    con_Ins_PRS.Open();
                                    SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_PRS, con_Ins_PRS);
                                    Cmd_Remp_AMS.ExecuteNonQuery();
                                    con_Ins_PRS.Close();
                                }
                                if (OPG1 && OPG2)
                                {
                                    while (Reader_OPG1.Read() && Reader_OPG2.Read())
                                    {
                                        if (Reader_OPG1[0].ToString() == "")
                                        {
                                            OPGId[j] = Reader_OPG2[0].ToString();
                                            OPGV[j] = Reader_OPG2[1].ToString();
                                            OPGId2[j] = Reader_OPG1[0].ToString();
                                            OPGV2[j] = Reader_OPG1[1].ToString();
                                        }
                                        else
                                        {
                                            OPGId[j] = Reader_OPG1[0].ToString();
                                            OPGV[j] = Reader_OPG1[1].ToString();
                                            OPGId2[j] = Reader_OPG2[0].ToString();
                                            OPGV2[j] = Reader_OPG2[1].ToString();
                                        }
                                        //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_I] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "')";
                                        Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_I] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2], [BundleID1], [BundleID2], [OPGID3], [BundleID3], [OPGID4], [BundleID4], [OPGID5], [BundleID5], [OPGID6], [BundleID6], [ShipToLocID], [HPInvoiceNo], [EndUserID], [PartPurchPrice], [PartPurchPriceCC], [PartRequestedRebateAmount], [PartnerComment], [PartnerReportedCBN], [PartnerReference], [PartnerInterTransID], [IsDropShip], [CustChannelPurchId], [PurchaseAgreement], [ReporterPurchOrderID], [SuppliesTrackID], [IntercompanyFlag], [ProdSerialIdHP]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "', '" + Insertion_Espaces_PRS_15(BundleID1) + "', '" + Insertion_Espaces_PRS_15(BundleID2) + "', '" + Insertion_Espaces_PRS_10(OPGID3) + "', '" + Insertion_Espaces_PRS_15(BundleID3) + "', '" + Insertion_Espaces_PRS_10(OPGID4) + "', '" + Insertion_Espaces_PRS_15(BundleID4) + "', '" + Insertion_Espaces_PRS_10(OPGID5) + "', '" + Insertion_Espaces_PRS_15(BundleID5) + "', '" + Insertion_Espaces_PRS_10(OPGID6) + "', '" + Insertion_Espaces_PRS_15(BundleID6) + "', '" + Insertion_Espaces_PRS_17(ShipToLocID) + "', '" + Insertion_Espaces_PRS_30(HPInvoiceNo) + "', '" + Insertion_Espaces_PRS_17(EndUserID) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3(PartPurchPriceCC) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_100(PartnerComment) + "', '" + Insertion_Espaces_PRS_15(PartnerReportedCBN) + "', '" + Insertion_Espaces_PRS_30(PartnerReference) + "', '" + Insertion_Espaces_PRS_40(PartnerInterTransID) + "', '" + Insertion_Espaces_PRS_6(IsDropShip) + "', '" + Insertion_Espaces_PRS_30(CustChannelPurchId) + "', '" + Insertion_Espaces_PRS_15(PurchaseAgreement) + "', '" + Insertion_Espaces_PRS_30(ReporterPurchOrderID) + "', '" + Insertion_Espaces_PRS_50(SuppliesTrackID) + "', '" + Insertion_Espaces_PRS_10(IntercompanyFlag) + "', '" + Insertion_Espaces_PRS_1600(ProdSerialIdHP) + "')";
                                        con_Ins_PRS.Open();
                                        SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_PRS, con_Ins_PRS);
                                        Cmd_Remp_AMS.ExecuteNonQuery();
                                        con_Ins_PRS.Close();
                                    }
                                }
                                if (OPG1 && !OPG2)
                                {
                                    while (Reader_OPG1.Read())
                                    {
                                        //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_I] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(Reader_OPG1[0].ToString()) + "','" + Insertion_Espaces_PRS_3(Reader_OPG1[1].ToString()) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "')";
                                        Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_I] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2], [BundleID1], [BundleID2], [OPGID3], [BundleID3], [OPGID4], [BundleID4], [OPGID5], [BundleID5], [OPGID6], [BundleID6], [ShipToLocID], [HPInvoiceNo], [EndUserID], [PartPurchPrice], [PartPurchPriceCC], [PartRequestedRebateAmount], [PartnerComment], [PartnerReportedCBN], [PartnerReference], [PartnerInterTransID], [IsDropShip], [CustChannelPurchId], [PurchaseAgreement], [ReporterPurchOrderID], [SuppliesTrackID], [IntercompanyFlag], [ProdSerialIdHP]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(Reader_OPG1[0].ToString()) + "','" + Insertion_Espaces_PRS_3(Reader_OPG1[1].ToString()) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "', '" + Insertion_Espaces_PRS_15(BundleID1) + "', '" + Insertion_Espaces_PRS_15(BundleID2) + "', '" + Insertion_Espaces_PRS_10(OPGID3) + "', '" + Insertion_Espaces_PRS_15(BundleID3) + "', '" + Insertion_Espaces_PRS_10(OPGID4) + "', '" + Insertion_Espaces_PRS_15(BundleID4) + "', '" + Insertion_Espaces_PRS_10(OPGID5) + "', '" + Insertion_Espaces_PRS_15(BundleID5) + "', '" + Insertion_Espaces_PRS_10(OPGID6) + "', '" + Insertion_Espaces_PRS_15(BundleID6) + "', '" + Insertion_Espaces_PRS_17(ShipToLocID) + "', '" + Insertion_Espaces_PRS_30(HPInvoiceNo) + "', '" + Insertion_Espaces_PRS_17(EndUserID) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3(PartPurchPriceCC) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_100(PartnerComment) + "', '" + Insertion_Espaces_PRS_15(PartnerReportedCBN) + "', '" + Insertion_Espaces_PRS_30(PartnerReference) + "', '" + Insertion_Espaces_PRS_40(PartnerInterTransID) + "', '" + Insertion_Espaces_PRS_6(IsDropShip) + "', '" + Insertion_Espaces_PRS_30(CustChannelPurchId) + "', '" + Insertion_Espaces_PRS_15(PurchaseAgreement) + "', '" + Insertion_Espaces_PRS_30(ReporterPurchOrderID) + "', '" + Insertion_Espaces_PRS_50(SuppliesTrackID) + "', '" + Insertion_Espaces_PRS_10(IntercompanyFlag) + "', '" + Insertion_Espaces_PRS_1600(ProdSerialIdHP) + "')";
                                        con_Ins_PRS.Open();
                                        SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_PRS, con_Ins_PRS);
                                        Cmd_Remp_AMS.ExecuteNonQuery();
                                        con_Ins_PRS.Close();
                                    }
                                }
                                if (!OPG1 && OPG2)
                                {
                                    while (Reader_OPG2.Read())
                                    {
                                        //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_I] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(Reader_OPG2[0].ToString()) + "','" + Insertion_Espaces_PRS_3(Reader_OPG2[1].ToString()) + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "')";
                                        Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_I] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2], [BundleID1], [BundleID2], [OPGID3], [BundleID3], [OPGID4], [BundleID4], [OPGID5], [BundleID5], [OPGID6], [BundleID6], [ShipToLocID], [HPInvoiceNo], [EndUserID], [PartPurchPrice], [PartPurchPriceCC], [PartRequestedRebateAmount], [PartnerComment], [PartnerReportedCBN], [PartnerReference], [PartnerInterTransID], [IsDropShip], [CustChannelPurchId], [PurchaseAgreement], [ReporterPurchOrderID], [SuppliesTrackID], [IntercompanyFlag], [ProdSerialIdHP]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(Reader_OPG2[0].ToString()) + "','" + Insertion_Espaces_PRS_3(Reader_OPG2[1].ToString()) + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "', '" + Insertion_Espaces_PRS_15(BundleID1) + "', '" + Insertion_Espaces_PRS_15(BundleID2) + "', '" + Insertion_Espaces_PRS_10(OPGID3) + "', '" + Insertion_Espaces_PRS_15(BundleID3) + "', '" + Insertion_Espaces_PRS_10(OPGID4) + "', '" + Insertion_Espaces_PRS_15(BundleID4) + "', '" + Insertion_Espaces_PRS_10(OPGID5) + "', '" + Insertion_Espaces_PRS_15(BundleID5) + "', '" + Insertion_Espaces_PRS_10(OPGID6) + "', '" + Insertion_Espaces_PRS_15(BundleID6) + "', '" + Insertion_Espaces_PRS_17(ShipToLocID) + "', '" + Insertion_Espaces_PRS_30(HPInvoiceNo) + "', '" + Insertion_Espaces_PRS_17(EndUserID) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3(PartPurchPriceCC) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_100(PartnerComment) + "', '" + Insertion_Espaces_PRS_15(PartnerReportedCBN) + "', '" + Insertion_Espaces_PRS_30(PartnerReference) + "', '" + Insertion_Espaces_PRS_40(PartnerInterTransID) + "', '" + Insertion_Espaces_PRS_6(IsDropShip) + "', '" + Insertion_Espaces_PRS_30(CustChannelPurchId) + "', '" + Insertion_Espaces_PRS_15(PurchaseAgreement) + "', '" + Insertion_Espaces_PRS_30(ReporterPurchOrderID) + "', '" + Insertion_Espaces_PRS_50(SuppliesTrackID) + "', '" + Insertion_Espaces_PRS_10(IntercompanyFlag) + "', '" + Insertion_Espaces_PRS_1600(ProdSerialIdHP) + "')";
                                        con_Ins_PRS.Open();
                                        SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_PRS, con_Ins_PRS);
                                        Cmd_Remp_AMS.ExecuteNonQuery();
                                        con_Ins_PRS.Close();
                                    }
                                }
                                Reader_OPG1.Close();
                                Reader_OPG2.Close();
                                con_OPG.Close();
                                con_OPG2.Close();
                            }
                            /* Remplissage SNT */
                            if (SN[j] == 1 && codePromo[j] != "")
                            {
                                Remplissage_SNT(codeHP, BL[j], NL_BL[j]);
                            }
                        }



                    }
                    /* SI la liste des clients exclus n'est pas vide on exécute la requête total */
                    else
                    {
                        /*string myquery =
    "SELECT S.[N° donneur d'ordre], S.[N°], CONVERT(int,ISNULL(SUM(S.Quantité),0)) AS Quantité, MAX(S.[N° lot]) AS [N° lot], S.[N° document] = CASE MAX(S.[N° lot]) WHEN '' THEN '' ELSE MAX(S.[N° document]) END, MAX(S.[Date comptabilisation]) AS InvoiceDate, ShipmentDate = CASE MAX(S.[N° lot]) WHEN '' THEN '' ELSE CONVERT(char, MAX(S.[Date comptabilisation])) END FROM dbo.___Sold_CM_BL S INNER JOIN dbo.___Product P ON S.[N°] = P.[N°] WHERE (((P.[Code Famille] IN (" + H + ")) AND (S.[N° donneur d'ordre] NOT IN (" + L + ")) AND (P.[N°] NOT IN (" + M + ")) AND (S.[Date comptabilisation] BETWEEN '" + dateTimePicker1.Text + "' AND '" + dateTimePicker2.Text + "') AND (S.[Code Magasin] NOT IN (" + EX + ") )))  GROUP BY S.[N° donneur d'ordre], S.[N°] Having SUM(S.Quantité)<> 0 ORDER BY S.[N° donneur d'ordre]";
                        /**** ZHA 29/06/2015 : OPG HP
                         */
                        string myquery =
    "SELECT S.[N° donneur d'ordre], S.[N°], CONVERT(int,ISNULL(SUM(S.Quantité),0)) AS Quantité, MAX(S.[N° lot]) AS [N° lot], S.[N° document], CONVERT(varchar(10),MAX(S.[Date comptabilisation]),126) AS InvoiceDate, CONVERT(varchar(10),MAX(S.[Date comptabilisation]),126) AS ShipmentDate ,S.[OPG Id],S.[OPG Version],S.[Secondary OPG Id],S.[Secondary OPG Version],S.[N° BL],S.[N° ligne BL],S.[Serial No_ Tracking],S.[Promotion code],S.[Entry No_] FROM dbo.___Sold_CM_BL_OPG S INNER JOIN dbo.___Product P ON S.[N°] = P.[N°] WHERE (((P.[Code Famille] IN (" + H + ")) " + Get_CodeSFamille(codeHP) + " AND (S.[Date comptabilisation] BETWEEN '" + dateTimePicker1.Text + "' AND '" + dateTimePicker2.Text + "') " + CLT + PRD + MAG + ")) GROUP BY S.[N° donneur d'ordre], S.[N° document], S.[N°],[OPG Id],[OPG Version],[Secondary OPG Id],[Secondary OPG Version],S.[N° BL],S.[N° ligne BL],S.[Serial No_ Tracking],S.[Promotion code],S.[Entry No_] Having SUM(S.Quantité)<> 0 ORDER BY S.[N° donneur d'ordre]"; 
                        
                         
                        SqlCommand Cmd_HDR = new SqlCommand(myquery, con6_PRS);
                        Cmd_HDR.CommandTimeout = 90;
                        SqlDataReader Reader_PRS = Cmd_HDR.ExecuteReader();
                        /* */
                        DataTable PRS_DT = new DataTable();
                        PRS_DT.Load(Reader_PRS);
                        int Count_myquery = PRS_DT.Rows.Count;
                        Reader_PRS = Cmd_HDR.ExecuteReader();
                        /*  */
                        int Count = Count_myquery + 1;
                        string[] Z = new string[Count];
                        string[] Y = new string[Count];
                        string[] G = new string[Count];
                        string[] F = new string[Count];
                        string[] W = new string[Count];
                        string[] V = new string[Count];
                        string[] U = new string[Count];
                        /**
                         *** ZHA 29/06/2015 : HP OPG
                         */
                        string[] OPGId = new string[Count];
                        string[] OPGV = new string[Count];
                        string[] OPGId2 = new string[Count];
                        string[] OPGV2 = new string[Count];
                        string[] BL = new string[Count];
                        string[] NL_BL = new string[Count];
                        string[] codePromo = new string[Count];
                        int[] SN = new int[Count];
                        int[] EntryNo = new int[Count];
                        int b = 0;
                        while (Reader_PRS.Read())
                        {
                            Z[b] = Reader_PRS.GetString(0);
                            Y[b] = Reader_PRS.GetString(1);
                            G[b] = Reader_PRS[2].ToString();
                            F[b] = Reader_PRS.GetString(3);
                            W[b] = Reader_PRS.GetString(4);
                            V[b] = Reader_PRS[5].ToString();
                            U[b] = Reader_PRS.GetString(6);
                            /**
                            *** ZHA 29/06/2015 : HP OPG
                            */
                            OPGId[b] = Reader_PRS[7].ToString(); //.GetString(7);
                            OPGV[b] = Reader_PRS[8].ToString();
                            OPGId2[b] = Reader_PRS[9].ToString();
                            OPGV2[b] = Reader_PRS[10].ToString();
                            BL[b] = Reader_PRS[11].ToString();
                            NL_BL[b] = Reader_PRS[12].ToString();
                            SN[b] = int.Parse(Reader_PRS[13].ToString());
                            codePromo[b] = Reader_PRS[14].ToString();
                            EntryNo[b] = int.Parse(Reader_PRS[15].ToString());
                            b++;
                        }
                        Reader_PRS.Close();

                        string BundleID1 = " ";
                        string BundleID2 = " ";
                        string OPGID3 = " ";
                        string BundleID3 = " ";
                        string OPGID4 = " ";
                        string BundleID4 = " ";
                        string OPGID5 = " ";
                        string BundleID5 = " ";
                        string OPGID6 = " ";
                        string BundleID6 = " ";
                        
                        string HPInvoiceNo = " ";
                        string EndUserID = " ";
                        //string PartPurchPrice = " ";
                        string PartPurchPriceCC = " ";
                        //string PartRequestedRebateAmount = " ";
                        string PartnerComment = " ";
                        string PartnerReportedCBN = " ";
                        string PartnerReference = " ";
                        string PartnerInterTransID = " ";
                        string IsDropShip = " ";
                        string CustChannelPurchId = " ";
                        string PurchaseAgreement = " ";
                        string ReporterPurchOrderID = " ";
                        string SuppliesTrackID = " ";
                        string IntercompanyFlag = " ";
                        string ProdSerialIdHP = " ";
                            
                        for (int j = 0; j < Count - 1; j++)
                        {
                            string ShipToLocID = Z[j];
                            SqlConnection con_Ins_PRS = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                            //con_Ins_PRS.Open();
                            //String Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_I] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(U[j]) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "')";
                            /**
                             *** ZHA 29/06/2015 : HP OPG
                             */
                            String Query_INS_PRS = "";
                            if (EntryNo[j] == 0)
                            {
                                //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_I] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "')";
                                Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_I] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2], [BundleID1], [BundleID2], [OPGID3], [BundleID3], [OPGID4], [BundleID4], [OPGID5], [BundleID5], [OPGID6], [BundleID6], [ShipToLocID], [HPInvoiceNo], [EndUserID], [PartPurchPrice], [PartPurchPriceCC], [PartRequestedRebateAmount], [PartnerComment], [PartnerReportedCBN], [PartnerReference], [PartnerInterTransID], [IsDropShip], [CustChannelPurchId], [PurchaseAgreement], [ReporterPurchOrderID], [SuppliesTrackID], [IntercompanyFlag], [ProdSerialIdHP]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "', '" + Insertion_Espaces_PRS_15(BundleID1) + "', '" + Insertion_Espaces_PRS_15(BundleID2) + "', '" + Insertion_Espaces_PRS_10(OPGID3) + "', '" + Insertion_Espaces_PRS_15(BundleID3) + "', '" + Insertion_Espaces_PRS_10(OPGID4) + "', '" + Insertion_Espaces_PRS_15(BundleID4) + "', '" + Insertion_Espaces_PRS_10(OPGID5) + "', '" + Insertion_Espaces_PRS_15(BundleID5) + "', '" + Insertion_Espaces_PRS_10(OPGID6) + "', '" + Insertion_Espaces_PRS_15(BundleID6) + "', '" + Insertion_Espaces_PRS_17(ShipToLocID) + "', '" + Insertion_Espaces_PRS_30(HPInvoiceNo) + "', '" + Insertion_Espaces_PRS_17(EndUserID) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3(PartPurchPriceCC) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_100(PartnerComment) + "', '" + Insertion_Espaces_PRS_15(PartnerReportedCBN) + "', '" + Insertion_Espaces_PRS_30(PartnerReference) + "', '" + Insertion_Espaces_PRS_40(PartnerInterTransID) + "', '" + Insertion_Espaces_PRS_6(IsDropShip) + "', '" + Insertion_Espaces_PRS_30(CustChannelPurchId) + "', '" + Insertion_Espaces_PRS_15(PurchaseAgreement) + "', '" + Insertion_Espaces_PRS_30(ReporterPurchOrderID) + "', '" + Insertion_Espaces_PRS_50(SuppliesTrackID) + "', '" + Insertion_Espaces_PRS_10(IntercompanyFlag) + "', '" + Insertion_Espaces_PRS_1600(ProdSerialIdHP) + "')";
                                con_Ins_PRS.Open();
                                SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_PRS, con_Ins_PRS);
                                Cmd_Remp_AMS.ExecuteNonQuery();
                                con_Ins_PRS.Close();

                            }
                            else
                            {
                                SqlConnection con_OPG = new SqlConnection("Data Source=WAYBI;Initial Catalog=Disway;Integrated Security=True");
                                con_OPG.Open();
                                SqlConnection con_OPG2 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Disway;Integrated Security=True");
                                con_OPG2.Open();
                                /*********** OPG 1 ************/
                                string R_OPG1 = "select [OPG Id],[OPG Version] from __OPG_HP_PRS_Primary where ILE_No=" + EntryNo[j] + " and [Item No_]='" + Y[j] + "'";
                                SqlCommand Cmd_OPG1 = new SqlCommand(R_OPG1, con_OPG);
                                Cmd_OPG1.CommandTimeout = 90;
                                SqlDataReader Reader_OPG1 = Cmd_OPG1.ExecuteReader();
                                /* On récupère le nombre de lignes reourné par la requête*/
                                bool OPG1 = Reader_OPG1.HasRows;

                                /*********** OPG 2 ************/
                                string R_OPG2 = "select [OPG Id 2],[OPG Version 2] from __OPG_HP_PRS_SecondeV2 where ILE_No=" + EntryNo[j] + " and [Item No_]='" + Y[j] + "'";
                                SqlCommand Cmd_OPG2 = new SqlCommand(R_OPG2, con_OPG2);
                                Cmd_OPG2.CommandTimeout = 90;
                                SqlDataReader Reader_OPG2 = Cmd_OPG2.ExecuteReader();
                                /* On récupère le nombre de lignes reourné par la requête*/
                                bool OPG2 = Reader_OPG2.HasRows;

                                if (!OPG1 && !OPG2)
                                {
                                    //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_I] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "')";
                                    Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_I] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2], [BundleID1], [BundleID2], [OPGID3], [BundleID3], [OPGID4], [BundleID4], [OPGID5], [BundleID5], [OPGID6], [BundleID6], [ShipToLocID], [HPInvoiceNo], [EndUserID], [PartPurchPrice], [PartPurchPriceCC], [PartRequestedRebateAmount], [PartnerComment], [PartnerReportedCBN], [PartnerReference], [PartnerInterTransID], [IsDropShip], [CustChannelPurchId], [PurchaseAgreement], [ReporterPurchOrderID], [SuppliesTrackID], [IntercompanyFlag], [ProdSerialIdHP]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "', '" + Insertion_Espaces_PRS_15(BundleID1) + "', '" + Insertion_Espaces_PRS_15(BundleID2) + "', '" + Insertion_Espaces_PRS_10(OPGID3) + "', '" + Insertion_Espaces_PRS_15(BundleID3) + "', '" + Insertion_Espaces_PRS_10(OPGID4) + "', '" + Insertion_Espaces_PRS_15(BundleID4) + "', '" + Insertion_Espaces_PRS_10(OPGID5) + "', '" + Insertion_Espaces_PRS_15(BundleID5) + "', '" + Insertion_Espaces_PRS_10(OPGID6) + "', '" + Insertion_Espaces_PRS_15(BundleID6) + "', '" + Insertion_Espaces_PRS_17(ShipToLocID) + "', '" + Insertion_Espaces_PRS_30(HPInvoiceNo) + "', '" + Insertion_Espaces_PRS_17(EndUserID) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3(PartPurchPriceCC) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_100(PartnerComment) + "', '" + Insertion_Espaces_PRS_15(PartnerReportedCBN) + "', '" + Insertion_Espaces_PRS_30(PartnerReference) + "', '" + Insertion_Espaces_PRS_40(PartnerInterTransID) + "', '" + Insertion_Espaces_PRS_6(IsDropShip) + "', '" + Insertion_Espaces_PRS_30(CustChannelPurchId) + "', '" + Insertion_Espaces_PRS_15(PurchaseAgreement) + "', '" + Insertion_Espaces_PRS_30(ReporterPurchOrderID) + "', '" + Insertion_Espaces_PRS_50(SuppliesTrackID) + "', '" + Insertion_Espaces_PRS_10(IntercompanyFlag) + "', '" + Insertion_Espaces_PRS_1600(ProdSerialIdHP) + "')";
                                    con_Ins_PRS.Open();
                                    SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_PRS, con_Ins_PRS);
                                    Cmd_Remp_AMS.ExecuteNonQuery();
                                    con_Ins_PRS.Close();
                                }
                                if (OPG1 && OPG2)
                                {
                                    while (Reader_OPG1.Read() && Reader_OPG2.Read())
                                    {
                                        if (Reader_OPG1[0].ToString() == "")
                                        {
                                            OPGId[j] = Reader_OPG2[0].ToString();
                                            OPGV[j] = Reader_OPG2[1].ToString();
                                            OPGId2[j] = Reader_OPG1[0].ToString();
                                            OPGV2[j] = Reader_OPG1[1].ToString();
                                        }
                                        else
                                        {
                                            OPGId[j] = Reader_OPG1[0].ToString();
                                            OPGV[j] = Reader_OPG1[1].ToString();
                                            OPGId2[j] = Reader_OPG2[0].ToString();
                                            OPGV2[j] = Reader_OPG2[1].ToString();
                                        }
                                        //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_I] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "')";
                                        Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_I] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2], [BundleID1], [BundleID2], [OPGID3], [BundleID3], [OPGID4], [BundleID4], [OPGID5], [BundleID5], [OPGID6], [BundleID6], [ShipToLocID], [HPInvoiceNo], [EndUserID], [PartPurchPrice], [PartPurchPriceCC], [PartRequestedRebateAmount], [PartnerComment], [PartnerReportedCBN], [PartnerReference], [PartnerInterTransID], [IsDropShip], [CustChannelPurchId], [PurchaseAgreement], [ReporterPurchOrderID], [SuppliesTrackID], [IntercompanyFlag], [ProdSerialIdHP]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "', '" + Insertion_Espaces_PRS_15(BundleID1) + "', '" + Insertion_Espaces_PRS_15(BundleID2) + "', '" + Insertion_Espaces_PRS_10(OPGID3) + "', '" + Insertion_Espaces_PRS_15(BundleID3) + "', '" + Insertion_Espaces_PRS_10(OPGID4) + "', '" + Insertion_Espaces_PRS_15(BundleID4) + "', '" + Insertion_Espaces_PRS_10(OPGID5) + "', '" + Insertion_Espaces_PRS_15(BundleID5) + "', '" + Insertion_Espaces_PRS_10(OPGID6) + "', '" + Insertion_Espaces_PRS_15(BundleID6) + "', '" + Insertion_Espaces_PRS_17(ShipToLocID) + "', '" + Insertion_Espaces_PRS_30(HPInvoiceNo) + "', '" + Insertion_Espaces_PRS_17(EndUserID) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3(PartPurchPriceCC) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_100(PartnerComment) + "', '" + Insertion_Espaces_PRS_15(PartnerReportedCBN) + "', '" + Insertion_Espaces_PRS_30(PartnerReference) + "', '" + Insertion_Espaces_PRS_40(PartnerInterTransID) + "', '" + Insertion_Espaces_PRS_6(IsDropShip) + "', '" + Insertion_Espaces_PRS_30(CustChannelPurchId) + "', '" + Insertion_Espaces_PRS_15(PurchaseAgreement) + "', '" + Insertion_Espaces_PRS_30(ReporterPurchOrderID) + "', '" + Insertion_Espaces_PRS_50(SuppliesTrackID) + "', '" + Insertion_Espaces_PRS_10(IntercompanyFlag) + "', '" + Insertion_Espaces_PRS_1600(ProdSerialIdHP) + "')";
                                        con_Ins_PRS.Open();
                                        SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_PRS, con_Ins_PRS);
                                        Cmd_Remp_AMS.ExecuteNonQuery();
                                        con_Ins_PRS.Close();
                                    }

                                }
                                if (OPG1 && !OPG2)
                                {
                                    while (Reader_OPG1.Read())
                                    {
                                        //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_I] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(Reader_OPG1[0].ToString()) + "','" + Insertion_Espaces_PRS_3(Reader_OPG1[1].ToString()) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "')";
                                        Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_I] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2], [BundleID1], [BundleID2], [OPGID3], [BundleID3], [OPGID4], [BundleID4], [OPGID5], [BundleID5], [OPGID6], [BundleID6], [ShipToLocID], [HPInvoiceNo], [EndUserID], [PartPurchPrice], [PartPurchPriceCC], [PartRequestedRebateAmount], [PartnerComment], [PartnerReportedCBN], [PartnerReference], [PartnerInterTransID], [IsDropShip], [CustChannelPurchId], [PurchaseAgreement], [ReporterPurchOrderID], [SuppliesTrackID], [IntercompanyFlag], [ProdSerialIdHP]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(Reader_OPG1[0].ToString()) + "','" + Insertion_Espaces_PRS_3(Reader_OPG1[1].ToString()) + "','" + Insertion_Espaces_PRS_10(OPGId2[j]) + "','" + Insertion_Espaces_PRS_3(OPGV2[j]) + "', '" + Insertion_Espaces_PRS_15(BundleID1) + "', '" + Insertion_Espaces_PRS_15(BundleID2) + "', '" + Insertion_Espaces_PRS_10(OPGID3) + "', '" + Insertion_Espaces_PRS_15(BundleID3) + "', '" + Insertion_Espaces_PRS_10(OPGID4) + "', '" + Insertion_Espaces_PRS_15(BundleID4) + "', '" + Insertion_Espaces_PRS_10(OPGID5) + "', '" + Insertion_Espaces_PRS_15(BundleID5) + "', '" + Insertion_Espaces_PRS_10(OPGID6) + "', '" + Insertion_Espaces_PRS_15(BundleID6) + "', '" + Insertion_Espaces_PRS_17(ShipToLocID) + "', '" + Insertion_Espaces_PRS_30(HPInvoiceNo) + "', '" + Insertion_Espaces_PRS_17(EndUserID) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3(PartPurchPriceCC) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_100(PartnerComment) + "', '" + Insertion_Espaces_PRS_15(PartnerReportedCBN) + "', '" + Insertion_Espaces_PRS_30(PartnerReference) + "', '" + Insertion_Espaces_PRS_40(PartnerInterTransID) + "', '" + Insertion_Espaces_PRS_6(IsDropShip) + "', '" + Insertion_Espaces_PRS_30(CustChannelPurchId) + "', '" + Insertion_Espaces_PRS_15(PurchaseAgreement) + "', '" + Insertion_Espaces_PRS_30(ReporterPurchOrderID) + "', '" + Insertion_Espaces_PRS_50(SuppliesTrackID) + "', '" + Insertion_Espaces_PRS_10(IntercompanyFlag) + "', '" + Insertion_Espaces_PRS_1600(ProdSerialIdHP) + "')";
                                        con_Ins_PRS.Open();
                                        SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_PRS, con_Ins_PRS);
                                        Cmd_Remp_AMS.ExecuteNonQuery();
                                        con_Ins_PRS.Close();
                                    }
                                }
                                if (!OPG1 && OPG2)
                                {
                                    while (Reader_OPG2.Read())
                                    {
                                        //Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_I] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(Reader_OPG2[0].ToString()) + "','" + Insertion_Espaces_PRS_3(Reader_OPG2[1].ToString()) + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "')";
                                        Query_INS_PRS = " insert into [" + textBox5.Text + "PRS_I] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber], [DetailedSellout] ,[SpecialPrincingreference]  ,[InvoiceNumber] ,[InvoiceDate] ,[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2], [BundleID1], [BundleID2], [OPGID3], [BundleID3], [OPGID4], [BundleID4], [OPGID5], [BundleID5], [OPGID6], [BundleID6], [ShipToLocID], [HPInvoiceNo], [EndUserID], [PartPurchPrice], [PartPurchPriceCC], [PartRequestedRebateAmount], [PartnerComment], [PartnerReportedCBN], [PartnerReference], [PartnerInterTransID], [IsDropShip], [CustChannelPurchId], [PurchaseAgreement], [ReporterPurchOrderID], [SuppliesTrackID], [IntercompanyFlag], [ProdSerialIdHP]) VALUES ('" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_17(Z[j]) + "', '" + Insertion_Espaces_PRS_20(Y[j]) + "','" + Insertion_Espaces_PRS_20("") + "', '" + Insertion_Negatif(Insertion_Espaces_SIT_10(Insertion_ZERO(G[j]))) + "','" + Insertion_Espaces_PRS_6(F[j]) + "' ,'" + Insertion_Espaces_PRS_30(W[j]) + "', '" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "','" + Insertion_Espaces_PRS_8(V[j].Replace("-", "")) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3("") + "','" + Insertion_Espaces_PRS_10(Reader_OPG2[0].ToString()) + "','" + Insertion_Espaces_PRS_3(Reader_OPG2[1].ToString()) + "','" + Insertion_Espaces_PRS_10(OPGId[j]) + "','" + Insertion_Espaces_PRS_3(OPGV[j]) + "', '" + Insertion_Espaces_PRS_15(BundleID1) + "', '" + Insertion_Espaces_PRS_15(BundleID2) + "', '" + Insertion_Espaces_PRS_10(OPGID3) + "', '" + Insertion_Espaces_PRS_15(BundleID3) + "', '" + Insertion_Espaces_PRS_10(OPGID4) + "', '" + Insertion_Espaces_PRS_15(BundleID4) + "', '" + Insertion_Espaces_PRS_10(OPGID5) + "', '" + Insertion_Espaces_PRS_15(BundleID5) + "', '" + Insertion_Espaces_PRS_10(OPGID6) + "', '" + Insertion_Espaces_PRS_15(BundleID6) + "', '" + Insertion_Espaces_PRS_17(ShipToLocID) + "', '" + Insertion_Espaces_PRS_30(HPInvoiceNo) + "', '" + Insertion_Espaces_PRS_17(EndUserID) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_3(PartPurchPriceCC) + "', '" + Insertion_Espaces_PRS_17("") + "', '" + Insertion_Espaces_PRS_100(PartnerComment) + "', '" + Insertion_Espaces_PRS_15(PartnerReportedCBN) + "', '" + Insertion_Espaces_PRS_30(PartnerReference) + "', '" + Insertion_Espaces_PRS_40(PartnerInterTransID) + "', '" + Insertion_Espaces_PRS_6(IsDropShip) + "', '" + Insertion_Espaces_PRS_30(CustChannelPurchId) + "', '" + Insertion_Espaces_PRS_15(PurchaseAgreement) + "', '" + Insertion_Espaces_PRS_30(ReporterPurchOrderID) + "', '" + Insertion_Espaces_PRS_50(SuppliesTrackID) + "', '" + Insertion_Espaces_PRS_10(IntercompanyFlag) + "', '" + Insertion_Espaces_PRS_1600(ProdSerialIdHP) + "')";
                                        con_Ins_PRS.Open();
                                        SqlCommand Cmd_Remp_AMS = new SqlCommand(Query_INS_PRS, con_Ins_PRS);
                                        Cmd_Remp_AMS.ExecuteNonQuery();
                                        con_Ins_PRS.Close();
                                    }
                                }
                                Reader_OPG1.Close();
                                Reader_OPG2.Close();
                                con_OPG.Close();
                                con_OPG2.Close();
                            }
                            /* Remplissage SNT */
                            if (SN[j] == 1 && codePromo[j] != "")
                            {
                                Remplissage_SNT(codeHP, BL[j], NL_BL[j]);
                            }
                        }

                    }
                    con6_PRS.Close();
                    richTextBox4.Text = richTextBox4.Text + " / " + "OK PRS_I";
                }
                return true;
            }

            catch (Exception e)
            {
                if (codeHP == CodeHPE)  MessageBox.Show("Remplissage PRS_E Not OK! ");
                else MessageBox.Show("Remplissage PRS_I Not OK! ");
                SqlConnection con4 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                con4.Open();
                string myquery2 = " insert into [Incidents_HP] ([Numero Etat HP] ,[Incident]) values (" + textBox5.Text.ToString() + ",'" + e.Message.ToString().Replace("'", "") + "')";
                
                SqlCommand Incident = new SqlCommand(myquery2, con4);
                Incident.ExecuteNonQuery();
                con4.Close();
                if (codeHP == CodeHPE)
                {
                    /*Notification_Mail*/
                    SendEmail(" notifhp@disway.com", "notifhp@disway.com", "Alerte HPE", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de PRS_E: " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                else
                {
                    /*Notification_Mail*/
                    SendEmail(" notifhp@disway.com", "notifhp@disway.com", "Alerte HPI", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de PRS_I: " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                return false;
            }
        }

        /* Fonction de création de la table TRL_E */
        private bool Creation_TRL_Table(string codeHP)
        {
            try
            {
                if (codeHP == CodeHPE)
                {
                    SqlConnection con5_TRL = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                    con5_TRL.Open();
                    string myquery = "CREATE TABLE [" + textBox5.Text + "TRL_E] (NumberOfRecords Varchar (50) PRIMARY KEY)";
                    SqlCommand Cmd_HDR =
                        new SqlCommand(myquery, con5_TRL);
                    Cmd_HDR.ExecuteNonQuery();
                    con5_TRL.Close();
                    /* MessageBox.Show("OK TRL_E");*/
                    richTextBox3.Text = richTextBox3.Text + " / " + "OK TRL_E";
                }
                else
                {
                    SqlConnection con5_TRL = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                    con5_TRL.Open();
                    string myquery = "CREATE TABLE [" + textBox5.Text + "TRL_I] (NumberOfRecords Varchar (50) PRIMARY KEY)";
                    SqlCommand Cmd_HDR =
                        new SqlCommand(myquery, con5_TRL);
                    Cmd_HDR.ExecuteNonQuery();
                    con5_TRL.Close();
                    /* MessageBox.Show("OK TRL_I");*/
                    richTextBox3.Text = richTextBox3.Text + " / " + "OK TRL_I";
                }
                return true;
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
                SqlConnection con4 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                con4.Open();
                string myquery2 = " insert into [Incidents_HP] ([Numero Etat HP] ,[Incident]) values (" + textBox5.Text.ToString() + ",'" + e.Message.ToString().Replace("'", "") + "')";
                SqlCommand Incident = new SqlCommand(myquery2, con4);
                Incident.ExecuteNonQuery();
                con4.Close();
                if (codeHP == CodeHPE)
                {
                    /*Notification_Mail*/
                    SendEmail("brahim.bourquia@gmail.com", "Disway.ReportingHp@gmail.com", "Alerte reporting HPE", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de TRL_E: " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                else
                {
                    /*Notification_Mail*/
                    SendEmail("brahim.bourquia@gmail.com", "Disway.ReportingHp@gmail.com", "Alerte reporting HPI", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de TRL_I: " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                return false;

            }
        }

        /* Fonction de remplissage de la table TRL_E */
        private bool Remplissage_TRL(string codeHP)
        {
            try
            {
                if (codeHP == CodeHPE)
                {
                    SqlConnection Con_rmp_TRL = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                    Con_rmp_TRL.Open();
                    /*Count HDR_E */
                    SqlCommand cmd_Count_HDR = new SqlCommand("select count(*) FROM [" + textBox5.Text + "HDR_E] ", Con_rmp_TRL);
                    int Count_HDR = (int)cmd_Count_HDR.ExecuteScalar();

                    /*Count AMS_E */
                    SqlCommand cmd_Count_AMS = new SqlCommand("select count(*) FROM [" + textBox5.Text + "AMS_E] ", Con_rmp_TRL);
                    int Count_AMS = (int)cmd_Count_AMS.ExecuteScalar();

                    /*Count CUS_E */
                    SqlCommand cmd_Count_CUS = new SqlCommand("select count(*) FROM [" + textBox5.Text + "CUS_E] ", Con_rmp_TRL);
                    int Count_CUS = (int)cmd_Count_CUS.ExecuteScalar();

                    /*Count PRS_E */
                    SqlCommand cmd_Count_PRS = new SqlCommand("select count(*) FROM [" + textBox5.Text + "PRS_E] ", Con_rmp_TRL);
                    int Count_PRS = (int)cmd_Count_PRS.ExecuteScalar();


                    /*Count SIT_E */
                    SqlCommand cmd_Count_SIT = new SqlCommand("select count(*) FROM [" + textBox5.Text + "SIT_E] ", Con_rmp_TRL);

                    int Count_SIT = (int)cmd_Count_SIT.ExecuteScalar();
                    /* Le 1 est le compte de la table TRL_E elle meme qui est tjrs égale à 1 */
                    int Count_Total = 1 + Count_HDR + Count_AMS + Count_CUS + Count_PRS + Count_SIT;

                    SqlCommand cmd_Ins_TRL = new SqlCommand("insert into [" + textBox5.Text + "TRL_E] ([NumberOfRecords]) values ('" + Insertion_ZERO(Count_Total.ToString()) + "')", Con_rmp_TRL);
                    cmd_Ins_TRL.ExecuteNonQuery();
                    Con_rmp_TRL.Close();
                    richTextBox4.Text = richTextBox4.Text + "  /  " + "OK TRL_E";
                }
                else
                {
                    SqlConnection Con_rmp_TRL = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                    Con_rmp_TRL.Open();
                    /*Count HDR_E */
                    SqlCommand cmd_Count_HDR = new SqlCommand("select count(*) FROM [" + textBox5.Text + "HDR_I] ", Con_rmp_TRL);
                    int Count_HDR = (int)cmd_Count_HDR.ExecuteScalar();

                    /*Count AMS_E */
                    SqlCommand cmd_Count_AMS = new SqlCommand("select count(*) FROM [" + textBox5.Text + "AMS_I] ", Con_rmp_TRL);
                    int Count_AMS = (int)cmd_Count_AMS.ExecuteScalar();

                    /*Count CUS_E */
                    SqlCommand cmd_Count_CUS = new SqlCommand("select count(*) FROM [" + textBox5.Text + "CUS_I] ", Con_rmp_TRL);
                    int Count_CUS = (int)cmd_Count_CUS.ExecuteScalar();

                    /*Count PRS_E */
                    SqlCommand cmd_Count_PRS = new SqlCommand("select count(*) FROM [" + textBox5.Text + "PRS_I] ", Con_rmp_TRL);
                    int Count_PRS = (int)cmd_Count_PRS.ExecuteScalar();


                    /*Count SIT_E */
                    SqlCommand cmd_Count_SIT = new SqlCommand("select count(*) FROM [" + textBox5.Text + "SIT_I] ", Con_rmp_TRL);

                    int Count_SIT = (int)cmd_Count_SIT.ExecuteScalar();
                    /* Le 1 est le compte de la table TRL_E elle meme qui est tjrs égale à 1 */
                    int Count_Total = 1 + Count_HDR + Count_AMS + Count_CUS + Count_PRS + Count_SIT;

                    SqlCommand cmd_Ins_TRL = new SqlCommand("insert into [" + textBox5.Text + "TRL_I] ([NumberOfRecords]) values ('" + Insertion_ZERO(Count_Total.ToString()) + "')", Con_rmp_TRL);
                    cmd_Ins_TRL.ExecuteNonQuery();
                    Con_rmp_TRL.Close();
                    richTextBox4.Text = richTextBox4.Text + "  /  " + "OK TRL_I";
                }
                return true;
            }
            catch (Exception e)
            {
                if (codeHP == CodeHPE)
                {
                    MessageBox.Show("Remplissage TRL_E Not OK! ");
                    MessageBox.Show(e.Message.ToString());
                }
                else
                {
                    MessageBox.Show("Remplissage TRL_I Not OK! ");
                    MessageBox.Show(e.Message.ToString());
                }
                SqlConnection con4 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                con4.Open();
                string myquery2 = " insert into [Incidents_HP] ([Numero Etat HP] ,[Incident]) values (" + textBox5.Text.ToString() + ",'" + e.Message.ToString().Replace("'", "") + "')";
              
                SqlCommand Incident = new SqlCommand(myquery2, con4);
                Incident.ExecuteNonQuery();
                con4.Close();
                if (codeHP == CodeHPE)
                {
                    /*Notification_Mail*/
                    SendEmail(" notifhp@disway.com", "notifhp@disway.com", "Alerte HPE", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de TRL_E: " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                else
                {
                    /*Notification_Mail*/
                    SendEmail(" notifhp@disway.com", "notifhp@disway.com", "Alerte HPI", "Le report HP n'a pas été envoyé ! ********** Cause : Erreur au niveau de TRL_I: " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                return false;
            }
        }
        
   
        /*Fonction de création et de remplissage du fichier*/
        private bool Creation_Report_HP(string codeHP)
        {
            
            if(codeHP.Equals(CodeHPE))
            {
                this.Devise();
            }
           if(codeHP.Equals(CodeHPI))
           {
               this.Traitement_Num_Serie_PRS_I();
                 this.Traitement_Num_Serie();

           }

           this.supp();
            try
            {
                if (codeHP == CodeHPE)
                {
                    
                    SqlConnection Con_HP = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                    string Query_HP_HDR = "select * from [" + textBox5.Text + "HDR_E] ";
                    string Query_HP_CUS = "select * from [" + textBox5.Text + "CUS_E]";
                    string Query_HP_SIT = "select * from [" + textBox5.Text + "SIT_E]";
                    string Query_HP_AMS = "select * from [" + textBox5.Text + "AMS_E]";
                    string Query_HP_PRS = "select * from [" + textBox5.Text + "PRS_E] order by 3";
                    string Query_HP_TRL = "select * from [" + textBox5.Text + "TRL_E]";
                    SqlCommand Cmd_HDR = new SqlCommand(Query_HP_HDR, Con_HP);
                    SqlCommand Cmd_CUS = new SqlCommand(Query_HP_CUS, Con_HP);
                    SqlCommand Cmd_SIT = new SqlCommand(Query_HP_SIT, Con_HP);
                    SqlCommand Cmd_AMS = new SqlCommand(Query_HP_AMS, Con_HP);
                    SqlCommand Cmd_PRS = new SqlCommand(Query_HP_PRS, Con_HP);
                    SqlCommand Cmd_TRL = new SqlCommand(Query_HP_TRL, Con_HP);

                    FileStream TheFile = File.Create(@"\\wayvs\Reporting\iflashas2_matel_" + CodeHPE + "." + textBox5.Text + ".dat");
                    StreamWriter Writer = new StreamWriter(TheFile, Encoding.GetEncoding(1252));

                    string TheFileSN = @"\\wayvs\Reporting\SNT_" + DateTime.Now.ToShortDateString().Replace("/", "") + "_" + CodeHPE + "." + textBox5.Text;
                    Excel_workout(codeHP,TheFileSN);

                    /*StreamWriter Writer_CUS = new StreamWriter(TheFile); */
                    string Header = "HDR";
                    string HDR_E = Header.Trim();

                    using (Con_HP)
                    {
                        Con_HP.Open();
                        /* Ecriture du HDR_E */
                        using (SqlDataReader Reader_HP_HDR = Cmd_HDR.ExecuteReader())
                        using (Writer)
                        {
                            while (Reader_HP_HDR.Read())
                            {
                                Writer.WriteLine(HDR_E.ToUpper().ToString() + Reader_HP_HDR[0].ToString() + "00" + Reader_HP_HDR[1].ToString() + Reader_HP_HDR[2].ToString() + Reader_HP_HDR[3].ToString() + Reader_HP_HDR[4].ToString() + Reader_HP_HDR[5].ToString() + Reader_HP_HDR[6].ToString() + Reader_HP_HDR[7].ToString() + Reader_HP_HDR[8].ToString() + Reader_HP_HDR[9].ToString() + Reader_HP_HDR[10].ToString() + Reader_HP_HDR[11].ToString());
                            }

                            /* Ecriture du CUS_E parès HDR_E*/
                            if (!Reader_HP_HDR.Read())
                            {
                                Reader_HP_HDR.Close();
                                SqlDataReader Reader_HP_SIT = Cmd_SIT.ExecuteReader();
                                while (Reader_HP_SIT.Read())
                                {     /* si le reserved inventory est supérieur au total inventory on les rends égaux */
                                    /* if (int.Parse(Reader_HP_SIT[5].ToString()) > int.Parse(Reader_HP_SIT[4].ToString()))*/
                                    if (int.Parse(Insertion_Negatif(Reader_HP_SIT[5].ToString())) > int.Parse(Insertion_Negatif(Reader_HP_SIT[4].ToString())))
                                    {
                                        //Writer.WriteLine("SIT" + Reader_HP_SIT[0].ToString() + Reader_HP_SIT[1].ToString() + Reader_HP_SIT[2].ToString() + Reader_HP_SIT[3].ToString() + Insertion_Negatif(Reader_HP_SIT[4].ToString()) + Insertion_Negatif(Reader_HP_SIT[4].ToString()) + Reader_HP_SIT[6].ToString() + Reader_HP_SIT[7].ToString() + Reader_HP_SIT[8].ToString() + Reader_HP_SIT[9].ToString());
                                        //Writer.WriteLine("SIT" + Reader_HP_SIT[0].ToString() + Reader_HP_SIT[1].ToString() + Reader_HP_SIT[2].ToString() + Reader_HP_SIT[3].ToString() + Insertion_Negatif(Reader_HP_SIT[4].ToString()) + Insertion_Negatif(Reader_HP_SIT[4].ToString()) + Reader_HP_SIT[6].ToString() + Reader_HP_SIT[7].ToString() + Reader_HP_SIT[8].ToString() + Reader_HP_SIT[9].ToString() + Reader_HP_SIT[10].ToString());
                                        Writer.WriteLine("SIT" + Reader_HP_SIT[0].ToString() + Reader_HP_SIT[1].ToString() + Reader_HP_SIT[2].ToString() + Reader_HP_SIT[3].ToString() + Insertion_Negatif(Reader_HP_SIT[4].ToString()) + Insertion_Negatif(Reader_HP_SIT[4].ToString()) + Reader_HP_SIT[6].ToString() + Reader_HP_SIT[7].ToString() + Reader_HP_SIT[8].ToString() + Reader_HP_SIT[9].ToString() + Reader_HP_SIT[10].ToString() + Reader_HP_SIT[11].ToString() + Reader_HP_SIT[12].ToString() + Reader_HP_SIT[13].ToString());
                                    }
                                    else
                                    {
                                        //Writer.WriteLine("SIT" + Reader_HP_SIT[0].ToString() + Reader_HP_SIT[1].ToString() + Reader_HP_SIT[2].ToString() + Reader_HP_SIT[3].ToString() + Insertion_Negatif(Reader_HP_SIT[4].ToString()) + Insertion_Negatif(Reader_HP_SIT[5].ToString()) + Reader_HP_SIT[6].ToString() + Reader_HP_SIT[7].ToString() + Reader_HP_SIT[8].ToString() + Reader_HP_SIT[9].ToString());
                                        //Writer.WriteLine("SIT" + Reader_HP_SIT[0].ToString() + Reader_HP_SIT[1].ToString() + Reader_HP_SIT[2].ToString() + Reader_HP_SIT[3].ToString() + Insertion_Negatif(Reader_HP_SIT[4].ToString()) + Insertion_Negatif(Reader_HP_SIT[5].ToString()) + Reader_HP_SIT[6].ToString() + Reader_HP_SIT[7].ToString() + Reader_HP_SIT[8].ToString() + Reader_HP_SIT[9].ToString() + Reader_HP_SIT[10].ToString());
                                        Writer.WriteLine("SIT" + Reader_HP_SIT[0].ToString() + Reader_HP_SIT[1].ToString() + Reader_HP_SIT[2].ToString() + Reader_HP_SIT[3].ToString() + Insertion_Negatif(Reader_HP_SIT[4].ToString()) + Insertion_Negatif(Reader_HP_SIT[5].ToString()) + Reader_HP_SIT[6].ToString() + Reader_HP_SIT[7].ToString() + Reader_HP_SIT[8].ToString() + Reader_HP_SIT[9].ToString() + Reader_HP_SIT[10].ToString() + Reader_HP_SIT[11].ToString() + Reader_HP_SIT[12].ToString() + Reader_HP_SIT[13].ToString());
                                    }
                                }
                                /*  Reader_HP_CUS.Close(); */

                                /* Ecriture du AMS_E parès HDR_E*/
                                if (!Reader_HP_SIT.Read())
                                {
                                    Reader_HP_SIT.Close();
                                    SqlDataReader Reader_HP_CUS = Cmd_CUS.ExecuteReader();
                                    while (Reader_HP_CUS.Read())
                                    {
                                        //Writer.WriteLine("CUS" + Reader_HP_CUS[0].ToString() + Reader_HP_CUS[1].ToString() + Reader_HP_CUS[2].ToString().ToUpper() + Reader_HP_CUS[3].ToString().ToUpper() + Reader_HP_CUS[4].ToString() + Reader_HP_CUS[5].ToString().ToUpper() + Reader_HP_CUS[6].ToString() + Reader_HP_CUS[7].ToString() + Reader_HP_CUS[8].ToString());
                                        //Writer.WriteLine("CUS" + Reader_HP_CUS[0].ToString() + Reader_HP_CUS[1].ToString() + Reader_HP_CUS[2].ToString().ToUpper() + Reader_HP_CUS[3].ToString().ToUpper() + Reader_HP_CUS[4].ToString() + Reader_HP_CUS[5].ToString().ToUpper() + Reader_HP_CUS[6].ToString() + Reader_HP_CUS[7].ToString() + Reader_HP_CUS[8].ToString() + Reader_HP_CUS[9].ToString().ToUpper());
                                        Writer.WriteLine("CUS" + Reader_HP_CUS[0].ToString() + Reader_HP_CUS[1].ToString() + Reader_HP_CUS[2].ToString().ToUpper() + Reader_HP_CUS[3].ToString().ToUpper() + Reader_HP_CUS[4].ToString() + Reader_HP_CUS[5].ToString().ToUpper() + Reader_HP_CUS[6].ToString() + Reader_HP_CUS[7].ToString() + Reader_HP_CUS[8].ToString() + Reader_HP_CUS[9].ToString().ToString());
                                    }
                                    /*   Reader_HP_AMS.Close(); */
                                    /* Ecriture du PRS_E après AMS_E*/
                                    if (!Reader_HP_CUS.Read())
                                    {
                                        Reader_HP_CUS.Close();
                                        SqlDataReader Reader_HP_PRS = Cmd_PRS.ExecuteReader();
                                        /************
                                         * ZHA HP-OPG 15/07/2015
                                         * Tester l'intégration de HP OPG dans le report
                                         * ***********/
                                        if (checkBox2.Checked)
                                        {
                                            while (Reader_HP_PRS.Read())
                                            {
                                                //Writer.WriteLine("PRS" + Reader_HP_PRS[0].ToString() + Reader_HP_PRS[1].ToString() + Reader_HP_PRS[2].ToString() + Reader_HP_PRS[3].ToString() + Reader_HP_PRS[4].ToString() + Reader_HP_PRS[5].ToString() + Reader_HP_PRS[6].ToString() + Reader_HP_PRS[7].ToString() + Reader_HP_PRS[8].ToString() + "00000000000000.00" + Reader_HP_PRS[10].ToString() +  Reader_HP_PRS[11].ToString() +  Reader_HP_PRS[12].ToString() +  Reader_HP_PRS[13].ToString() + Reader_HP_PRS[14].ToString());
                                                //Writer.WriteLine("PRS" + Reader_HP_PRS[0].ToString() + Reader_HP_PRS[1].ToString() + Reader_HP_PRS[2].ToString() + Reader_HP_PRS[3].ToString() + Reader_HP_PRS[4].ToString() + Reader_HP_PRS[5].ToString() + Reader_HP_PRS[6].ToString() + Reader_HP_PRS[7].ToString() + Reader_HP_PRS[8].ToString() + "00000000000000.00" + Reader_HP_PRS[10].ToString() + Reader_HP_PRS[11].ToString() + Reader_HP_PRS[12].ToString() + Reader_HP_PRS[13].ToString() + Reader_HP_PRS[14].ToString() + Reader_HP_PRS[15].ToString() + Reader_HP_PRS[16].ToString() + Reader_HP_PRS[17].ToString() + Reader_HP_PRS[18].ToString() + Reader_HP_PRS[19].ToString() + Reader_HP_PRS[20].ToString() + Reader_HP_PRS[21].ToString() + Reader_HP_PRS[22].ToString() + Reader_HP_PRS[23].ToString() + Reader_HP_PRS[24].ToString() + Reader_HP_PRS[25].ToString() + Reader_HP_PRS[26].ToString() + Reader_HP_PRS[27].ToString() + Reader_HP_PRS[28].ToString() + Reader_HP_PRS[29].ToString() + Reader_HP_PRS[30].ToString() + Reader_HP_PRS[31].ToString() + Reader_HP_PRS[32].ToString() + Reader_HP_PRS[33].ToString() + Reader_HP_PRS[34].ToString() + Reader_HP_PRS[35].ToString() + Reader_HP_PRS[36].ToString() + Reader_HP_PRS[37].ToString() + Reader_HP_PRS[38].ToString() + Reader_HP_PRS[39].ToString() + Reader_HP_PRS[40].ToString() + Reader_HP_PRS[41].ToString());
                                                Writer.WriteLine("PRS" + Reader_HP_PRS[0].ToString() + Reader_HP_PRS[1].ToString() + Reader_HP_PRS[2].ToString() + Reader_HP_PRS[3].ToString() + Reader_HP_PRS[4].ToString() + Reader_HP_PRS[5].ToString() + Reader_HP_PRS[6].ToString() + Reader_HP_PRS[7].ToString() + Reader_HP_PRS[8].ToString() + "00000000000000.00" + Reader_HP_PRS[10].ToString() + Reader_HP_PRS[11].ToString() + Reader_HP_PRS[12].ToString() + Reader_HP_PRS[13].ToString() + Reader_HP_PRS[14].ToString() + Reader_HP_PRS[15].ToString() + Reader_HP_PRS[16].ToString() + Reader_HP_PRS[17].ToString() + Reader_HP_PRS[18].ToString() + Reader_HP_PRS[19].ToString() + Reader_HP_PRS[20].ToString() + Reader_HP_PRS[21].ToString() + Reader_HP_PRS[22].ToString() + Reader_HP_PRS[23].ToString() + Reader_HP_PRS[24].ToString() + Reader_HP_PRS[25].ToString() + Reader_HP_PRS[26].ToString() + Reader_HP_PRS[27].ToString() + Reader_HP_PRS[28].ToString() + Reader_HP_PRS[29].ToString() + Reader_HP_PRS[30].ToString() + Reader_HP_PRS[31].ToString() + Reader_HP_PRS[32].ToString() + Reader_HP_PRS[33].ToString() + Reader_HP_PRS[34].ToString());
                                            }
                                        }
                                        else
                                        {
                                            while (Reader_HP_PRS.Read())
                                            {
                                                //Writer.WriteLine("PRS" + Reader_HP_PRS[0].ToString() + Reader_HP_PRS[1].ToString() + Reader_HP_PRS[2].ToString() + Reader_HP_PRS[3].ToString() + Reader_HP_PRS[4].ToString() + Reader_HP_PRS[5].ToString() + Reader_HP_PRS[6].ToString() + Reader_HP_PRS[7].ToString() + Reader_HP_PRS[8].ToString() + "00000000000000.00" + Reader_HP_PRS[10].ToString() + "                          ");
                                                //Writer.WriteLine("PRS" + Reader_HP_PRS[0].ToString() + Reader_HP_PRS[1].ToString() + Reader_HP_PRS[2].ToString() + Reader_HP_PRS[3].ToString() + Reader_HP_PRS[4].ToString() + Reader_HP_PRS[5].ToString() + Reader_HP_PRS[6].ToString() + Reader_HP_PRS[7].ToString() + Reader_HP_PRS[8].ToString() + "00000000000000.00" + Reader_HP_PRS[10].ToString() + "                          " + Reader_HP_PRS[15].ToString() + Reader_HP_PRS[16].ToString() + Reader_HP_PRS[17].ToString() + Reader_HP_PRS[18].ToString() + Reader_HP_PRS[19].ToString() + Reader_HP_PRS[20].ToString() + Reader_HP_PRS[21].ToString() + Reader_HP_PRS[22].ToString() + Reader_HP_PRS[23].ToString() + Reader_HP_PRS[24].ToString() + Reader_HP_PRS[25].ToString() + Reader_HP_PRS[26].ToString() + Reader_HP_PRS[27].ToString() + Reader_HP_PRS[28].ToString() + Reader_HP_PRS[29].ToString() + Reader_HP_PRS[30].ToString() + Reader_HP_PRS[31].ToString() + Reader_HP_PRS[32].ToString() + Reader_HP_PRS[33].ToString() + Reader_HP_PRS[34].ToString() + Reader_HP_PRS[35].ToString() + Reader_HP_PRS[36].ToString() + Reader_HP_PRS[37].ToString() + Reader_HP_PRS[38].ToString() + Reader_HP_PRS[39].ToString() + Reader_HP_PRS[40].ToString() + Reader_HP_PRS[41].ToString());
                                                Writer.WriteLine("PRS" + Reader_HP_PRS[0].ToString() + Reader_HP_PRS[1].ToString() + Reader_HP_PRS[2].ToString() + Reader_HP_PRS[3].ToString() + Reader_HP_PRS[4].ToString() + Reader_HP_PRS[5].ToString() + Reader_HP_PRS[6].ToString() + Reader_HP_PRS[7].ToString() + Reader_HP_PRS[8].ToString() + "00000000000000.00" + Reader_HP_PRS[10].ToString() + "                          " + Reader_HP_PRS[15].ToString() + Reader_HP_PRS[16].ToString() + Reader_HP_PRS[17].ToString() + Reader_HP_PRS[18].ToString() + Reader_HP_PRS[19].ToString() + Reader_HP_PRS[20].ToString() + Reader_HP_PRS[21].ToString() + Reader_HP_PRS[22].ToString() + Reader_HP_PRS[23].ToString() + Reader_HP_PRS[24].ToString() + Reader_HP_PRS[25].ToString() + Reader_HP_PRS[26].ToString() + Reader_HP_PRS[27].ToString() + Reader_HP_PRS[28].ToString() + Reader_HP_PRS[29].ToString() + Reader_HP_PRS[30].ToString() + Reader_HP_PRS[31].ToString() + Reader_HP_PRS[32].ToString() + Reader_HP_PRS[33].ToString() + Reader_HP_PRS[34].ToString());
                                            }
                                        }
                                        /*  Reader_HP_PRS.Close(); */

                                        /* Ecriture du SIT_E après PRS_E*/
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
                        
                    }
                }
                else
                {
                   
                    /*Remplissage du fichier dat */
                    SqlConnection Con_HP = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                    string Query_HP_HDR = "select * from [" + textBox5.Text + "HDR_I] ";
                    string Query_HP_CUS = "select * from [" + textBox5.Text + "CUS_I]";
                    string Query_HP_SIT = "select * from [" + textBox5.Text + "SIT_I]";
                    string Query_HP_AMS = "select * from [" + textBox5.Text + "AMS_I]";
                    string Query_HP_PRS = "select * from [" + textBox5.Text + "PRS_I_SN] order by 3";
                    string Query_HP_TRL = "select * from [" + textBox5.Text + "TRL_I]";
                    SqlCommand Cmd_HDR = new SqlCommand(Query_HP_HDR, Con_HP);
                    SqlCommand Cmd_CUS = new SqlCommand(Query_HP_CUS, Con_HP);
                    SqlCommand Cmd_SIT = new SqlCommand(Query_HP_SIT, Con_HP);
                    SqlCommand Cmd_AMS = new SqlCommand(Query_HP_AMS, Con_HP);
                    SqlCommand Cmd_PRS = new SqlCommand(Query_HP_PRS, Con_HP);
                    SqlCommand Cmd_TRL = new SqlCommand(Query_HP_TRL, Con_HP);

                    FileStream TheFile = File.Create(@"\\wayvs\Reporting\iflashas2_matel_" + CodeHPI + "." + textBox5.Text + ".dat");
                    StreamWriter Writer = new StreamWriter(TheFile, Encoding.GetEncoding(1252));

                    string TheFileSN = @"\\wayvs\Reporting\SNT_" + DateTime.Now.ToShortDateString().Replace("/", "") + "_" + CodeHPI + "." + textBox5.Text;
                    Excel_workout(codeHP,TheFileSN);

                    /*StreamWriter Writer_CUS = new StreamWriter(TheFile); */
                    string Header = "HDR";
                    string HDR_I = Header.Trim();

                    using (Con_HP)
                    {
                        Con_HP.Open();
                        /* Ecriture du HDR_E */
                        using (SqlDataReader Reader_HP_HDR = Cmd_HDR.ExecuteReader())
                        using (Writer)
                        {
                            while (Reader_HP_HDR.Read())
                            {
                                Writer.WriteLine(HDR_I.ToUpper().ToString() + Reader_HP_HDR[0].ToString() + "00" + Reader_HP_HDR[1].ToString() + Reader_HP_HDR[2].ToString() + Reader_HP_HDR[3].ToString() + Reader_HP_HDR[4].ToString() + Reader_HP_HDR[5].ToString() + Reader_HP_HDR[6].ToString() + Reader_HP_HDR[7].ToString() + Reader_HP_HDR[8].ToString() + Reader_HP_HDR[9].ToString() + Reader_HP_HDR[10].ToString() + Reader_HP_HDR[11].ToString());
                            }

                            /* Ecriture du CUS_E parès HDR_I*/
                            if (!Reader_HP_HDR.Read())
                            {
                                Reader_HP_HDR.Close();
                                SqlDataReader Reader_HP_SIT = Cmd_SIT.ExecuteReader();
                                while (Reader_HP_SIT.Read())
                                {     /* si le reserved inventory est supérieur au total inventory on les rends égaux */
                                    /* if (int.Parse(Reader_HP_SIT[5].ToString()) > int.Parse(Reader_HP_SIT[4].ToString()))*/
                                    if (int.Parse(Insertion_Negatif(Reader_HP_SIT[5].ToString())) > int.Parse(Insertion_Negatif(Reader_HP_SIT[4].ToString())))
                                    {
                                        /*-------------- Aurora 12/10/16 LHM------------------*/
                                        //Writer.WriteLine("SIT" + Reader_HP_SIT[0].ToString() + Reader_HP_SIT[1].ToString() + Reader_HP_SIT[2].ToString() + Reader_HP_SIT[3].ToString() + Insertion_Negatif(Reader_HP_SIT[4].ToString()) + Insertion_Negatif(Reader_HP_SIT[4].ToString()) + Reader_HP_SIT[6].ToString() + Reader_HP_SIT[7].ToString() + Reader_HP_SIT[8].ToString() + Reader_HP_SIT[9].ToString());
                                        Writer.WriteLine("SIT" + Reader_HP_SIT[0].ToString() + Reader_HP_SIT[1].ToString() + Reader_HP_SIT[2].ToString() + Reader_HP_SIT[3].ToString() + Insertion_Negatif(Reader_HP_SIT[4].ToString()) + Insertion_Negatif(Reader_HP_SIT[4].ToString()) + Reader_HP_SIT[6].ToString() + Reader_HP_SIT[7].ToString() + Reader_HP_SIT[8].ToString() + Reader_HP_SIT[9].ToString() + Reader_HP_SIT[10].ToString());
                                        /*-------------- Aurora 12/10/16 LHM------------------*/
                                    }
                                    else
                                    {
                                        /*-------------- Aurora 12/10/16 LHM------------------*/
                                        //Writer.WriteLine("SIT" + Reader_HP_SIT[0].ToString() + Reader_HP_SIT[1].ToString() + Reader_HP_SIT[2].ToString() + Reader_HP_SIT[3].ToString() + Insertion_Negatif(Reader_HP_SIT[4].ToString()) + Insertion_Negatif(Reader_HP_SIT[5].ToString()) + Reader_HP_SIT[6].ToString() + Reader_HP_SIT[7].ToString() + Reader_HP_SIT[8].ToString() + Reader_HP_SIT[9].ToString());
                                        Writer.WriteLine("SIT" + Reader_HP_SIT[0].ToString() + Reader_HP_SIT[1].ToString() + Reader_HP_SIT[2].ToString() + Reader_HP_SIT[3].ToString() + Insertion_Negatif(Reader_HP_SIT[4].ToString()) + Insertion_Negatif(Reader_HP_SIT[5].ToString()) + Reader_HP_SIT[6].ToString() + Reader_HP_SIT[7].ToString() + Reader_HP_SIT[8].ToString() + Reader_HP_SIT[9].ToString() + Reader_HP_SIT[10].ToString());
                                        /*-------------- Aurora 12/10/16 LHM------------------*/
                                    }
                                }
                                /*  Reader_HP_CUS.Close(); */

                                /* Ecriture du AMS_E parès HDR_E*/
                                if (!Reader_HP_SIT.Read())
                                {
                                    Reader_HP_SIT.Close();
                                    SqlDataReader Reader_HP_CUS = Cmd_CUS.ExecuteReader();
                                    while (Reader_HP_CUS.Read())
                                    {   //******************* AURORA LHM 17102016 ****************
                                       // Writer.WriteLine("CUS" + Reader_HP_CUS[0].ToString() + Reader_HP_CUS[1].ToString() + Reader_HP_CUS[2].ToString().ToUpper() + Reader_HP_CUS[3].ToString().ToUpper() + Reader_HP_CUS[4].ToString() + Reader_HP_CUS[5].ToString().ToUpper() + Reader_HP_CUS[6].ToString() + Reader_HP_CUS[7].ToString() + Reader_HP_CUS[8].ToString());
                                         Writer.WriteLine("CUS" + Reader_HP_CUS[0].ToString() + Reader_HP_CUS[1].ToString() + Reader_HP_CUS[2].ToString().ToUpper() + Reader_HP_CUS[3].ToString().ToUpper() + Reader_HP_CUS[4].ToString() + Reader_HP_CUS[5].ToString().ToUpper() + Reader_HP_CUS[6].ToString() + Reader_HP_CUS[7].ToString() + Reader_HP_CUS[8].ToString() + Reader_HP_CUS[9].ToString().ToUpper());
                                    }
                                    /*   Reader_HP_AMS.Close(); */
                                    /* Ecriture du PRS_E après AMS_I*/
                                    if (!Reader_HP_CUS.Read())
                                    {
                                        Reader_HP_CUS.Close();
                                        SqlDataReader Reader_HP_PRS = Cmd_PRS.ExecuteReader();
                                        /*AM 10/11/2015*/
                                         /* Tester l'intégration de HP OPG dans le report HPI
                                         * ***********/
                                        if (checkBox2.Checked)
                                        {
                                            while (Reader_HP_PRS.Read())
                                            {   //******************* AURORA LHM 17102016 ****************
                                                //Writer.WriteLine("PRS" + Reader_HP_PRS[0].ToString() + Reader_HP_PRS[1].ToString() + Reader_HP_PRS[2].ToString() + Reader_HP_PRS[3].ToString() + Reader_HP_PRS[4].ToString() + Reader_HP_PRS[5].ToString() + Reader_HP_PRS[6].ToString() + Reader_HP_PRS[7].ToString() + Reader_HP_PRS[8].ToString() + "00000000000000.00" + Reader_HP_PRS[10].ToString() +  Reader_HP_PRS[11].ToString() +  Reader_HP_PRS[12].ToString() +  Reader_HP_PRS[13].ToString() + Reader_HP_PRS[14].ToString());
                                                Writer.WriteLine("PRS" + Reader_HP_PRS[0].ToString() + Reader_HP_PRS[1].ToString() + Reader_HP_PRS[2].ToString() + Reader_HP_PRS[3].ToString() + Reader_HP_PRS[4].ToString() + Reader_HP_PRS[5].ToString() + Reader_HP_PRS[6].ToString() + Reader_HP_PRS[7].ToString() + Reader_HP_PRS[8].ToString() + "00000000000000.00" + Reader_HP_PRS[10].ToString() + Reader_HP_PRS[11].ToString() + Reader_HP_PRS[12].ToString() + Reader_HP_PRS[13].ToString() + Reader_HP_PRS[14].ToString() + Reader_HP_PRS[15].ToString() + Reader_HP_PRS[16].ToString() + Reader_HP_PRS[17].ToString() + Reader_HP_PRS[18].ToString() + Reader_HP_PRS[19].ToString() + Reader_HP_PRS[20].ToString() + Reader_HP_PRS[21].ToString() + Reader_HP_PRS[22].ToString() + Reader_HP_PRS[23].ToString() + Reader_HP_PRS[24].ToString() + Reader_HP_PRS[25].ToString() + Reader_HP_PRS[26].ToString() + Reader_HP_PRS[27].ToString() + Reader_HP_PRS[28].ToString() + Reader_HP_PRS[29].ToString() + Reader_HP_PRS[30].ToString() + Reader_HP_PRS[31].ToString() + Reader_HP_PRS[32].ToString() + Reader_HP_PRS[33].ToString() + Reader_HP_PRS[34].ToString() + Reader_HP_PRS[35].ToString() + Reader_HP_PRS[36].ToString() + Reader_HP_PRS[37].ToString() + Reader_HP_PRS[38].ToString() + Reader_HP_PRS[39].ToString() + Reader_HP_PRS[40].ToString() + Reader_HP_PRS[41].ToString());
                                            }
                                        }
                                        else
                                        {
                                            while (Reader_HP_PRS.Read())
                                            {
                                               // Writer.WriteLine("PRS" + Reader_HP_PRS[0].ToString() + Reader_HP_PRS[1].ToString() + Reader_HP_PRS[2].ToString() + Reader_HP_PRS[3].ToString() + Reader_HP_PRS[4].ToString() + Reader_HP_PRS[5].ToString() + Reader_HP_PRS[6].ToString() + Reader_HP_PRS[7].ToString() + Reader_HP_PRS[8].ToString() + "00000000000000.00" + Reader_HP_PRS[10].ToString() + "                          ");
                                                Writer.WriteLine("PRS" + Reader_HP_PRS[0].ToString() + Reader_HP_PRS[1].ToString() + Reader_HP_PRS[2].ToString() + Reader_HP_PRS[3].ToString() + Reader_HP_PRS[4].ToString() + Reader_HP_PRS[5].ToString() + Reader_HP_PRS[6].ToString() + Reader_HP_PRS[7].ToString() + Reader_HP_PRS[8].ToString() + "00000000000000.00" + Reader_HP_PRS[10].ToString() + "                          "  + Reader_HP_PRS[15].ToString() + Reader_HP_PRS[16].ToString() + Reader_HP_PRS[17].ToString() + Reader_HP_PRS[18].ToString() + Reader_HP_PRS[19].ToString() + Reader_HP_PRS[20].ToString() + Reader_HP_PRS[21].ToString() + Reader_HP_PRS[22].ToString() + Reader_HP_PRS[23].ToString() + Reader_HP_PRS[24].ToString() + Reader_HP_PRS[25].ToString() + Reader_HP_PRS[26].ToString() + Reader_HP_PRS[27].ToString() + Reader_HP_PRS[28].ToString() + Reader_HP_PRS[29].ToString() + Reader_HP_PRS[30].ToString() + Reader_HP_PRS[31].ToString() + Reader_HP_PRS[32].ToString() + Reader_HP_PRS[33].ToString() + Reader_HP_PRS[34].ToString() + Reader_HP_PRS[35].ToString() + Reader_HP_PRS[36].ToString() + Reader_HP_PRS[37].ToString() + Reader_HP_PRS[38].ToString() + Reader_HP_PRS[39].ToString() + Reader_HP_PRS[40].ToString() + Reader_HP_PRS[41].ToString());
                                            }
                                        }
                                        /* while (Reader_HP_PRS.Read())
                                         {  //******************* AURORA LHM 17102016 ****************
                                             //Writer.WriteLine("PRS" + Reader_HP_PRS[0].ToString() + Reader_HP_PRS[1].ToString() + Reader_HP_PRS[2].ToString() + Reader_HP_PRS[3].ToString() + Reader_HP_PRS[4].ToString() + Reader_HP_PRS[5].ToString() + Reader_HP_PRS[6].ToString() + Reader_HP_PRS[7].ToString() + Reader_HP_PRS[8].ToString() + "00000000000000.00" + Reader_HP_PRS[10].ToString());
                                         }
                                         * /
                                         /*  Reader_HP_PRS.Close(); */

                                        /* Ecriture du SIT_E après PRS_I*/
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

                    }
                }
                return true;
            }
            catch (Exception e)
            {
                if (codeHP == CodeHPE)
                {

                    MessageBox.Show("Etat HPE non crée !");
                    MessageBox.Show(e.Message.ToString());
                }
                else
                {
                    MessageBox.Show("Etat HPI non crée !");
                    MessageBox.Show(e.Message.ToString());
                }
                MessageBox.Show("Exception : " + e.Message + " ");
                SqlConnection con4 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                con4.Open();
                string myquery2 = "insert into [Incidents_HP] ([Numero Etat HP] ,[Incident]) values (" + textBox5.Text.ToString() + ",'" + e.Message.ToString().Replace("'", "") + "')";
                SqlCommand Incident = new SqlCommand(myquery2, con4);
                Incident.ExecuteNonQuery();
                con4.Close();
                if (codeHP == CodeHPE)
                {
                    /*Notification_Mail*/
                    SendEmail("notifhp@disway.com", "notifhp@disway.com", "Alerte HPE", "Le report HPE n'a pas été envoyé ! ********** Cause : Erreur au niveau du remplissage du fichier : " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                else
                {
                    /*Notification_Mail*/
                    SendEmail("notifhp@disway.com", "notifhp@disway.com", "Alerte HPI", "Le report HPI n'a pas été envoyé ! ********** Cause : Erreur au niveau du remplissage du fichier : " + e.Message.ToString() + "");
                    /*Fin notiification */
                }
                return false;
            }
        }
        /***
         * ZHA 18/06/15 Création fichier CSV pour les N° de série 
         ***/
        private void Excel_workout(string codeHP, string fileName)
        {
            if (codeHP == CodeHPE)
            {
                SqlConnection conn = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                conn.Open();
                string query = "select * from [" + textBox5.Text + "SNT_E] ";
                SqlCommand cmd = new SqlCommand(query, conn);
                //SqlDataReader dr = cmd.ExecuteReader();
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                int i = 0;
                int j = 0;

                DataTable dtExcelTable = new DataTable();

                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dtExcelTable);


                for (i = 0; i < dtExcelTable.Columns.Count; i++)
                {
                    xlWorkSheet.Cells[1, i + 1] = dtExcelTable.Columns[i].ColumnName;
                }
                if (checkBox2.Checked)
                {
                    for (i = 0; i < dtExcelTable.Rows.Count; i++)
                    {
                        for (j = 0; j < dtExcelTable.Columns.Count; j++)
                        {
                            xlWorkSheet.Cells[i + 2, j + 1] = dtExcelTable.Rows[i][j];
                        }
                    }
                }
                try
                {
                    xlWorkBook.SaveAs(fileName + ".xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkBook);
                    releaseObject(xlApp);

                    //MessageBox.Show("Excel crée d:\\" + fileName + ".xls");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error " + ex.Message);
                }
            }
            else
            {
                SqlConnection conn = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                conn.Open();
                string query = "select * from [" + textBox5.Text + "SNT_I] ";
                SqlCommand cmd = new SqlCommand(query, conn);
                //SqlDataReader dr = cmd.ExecuteReader();
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                int i = 0;
                int j = 0;

                DataTable dtExcelTable = new DataTable();

                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dtExcelTable);


                for (i = 0; i < dtExcelTable.Columns.Count; i++)
                {
                    xlWorkSheet.Cells[1, i + 1] = dtExcelTable.Columns[i].ColumnName;
                }
                if (checkBox2.Checked)
                {
                    for (i = 0; i < dtExcelTable.Rows.Count; i++)
                    {
                        for (j = 0; j < dtExcelTable.Columns.Count; j++)
                        {
                            xlWorkSheet.Cells[i + 2, j + 1] = dtExcelTable.Rows[i][j];
                        }
                    }
                }
                try
                {
                    xlWorkBook.SaveAs(fileName + ".xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkBook);
                    releaseObject(xlApp);

                    //MessageBox.Show("Excel crée d:\\" + fileName + ".xls");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error " + ex.Message);
                }
            }
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        /**
         * Fin
         **/
        /*Fonction de déplacement d'un fichier d'un réportoire vers un autre */
        private bool  Depos_HP(string CodeHP)
        {
            try
            {
                if (CodeHP == CodeHPE)
                {
                      File.Copy(@"\\wayvs\Reporting\iflashas2_matel_" + CodeHP + "." + textBox5.Text + ".dat", @"\\wayvs\Reporting\iflashas2_matel_bis." + CodeHP + "." + textBox5.Text + ".dat");
                      System.IO.File.Move(@"\\wayvs\Reporting\iflashas2_matel_bis." + CodeHP + "." + textBox5.Text + ".dat", @"\\wayedi\HPE\iflashas2_matel_" + CodeHP + "." + textBox5.Text + ".dat");
                      File.Delete(@"\\wayvs\Reporting\iflashas2_matel_bis." + CodeHP + "." + textBox5.Text + ".dat");
                    return true;
                }
                else
                {
                    File.Copy(@"\\wayvs\Reporting\iflashas2_matel_" + CodeHP + "." + textBox5.Text + ".dat", @"\\wayvs\Reporting\iflashas2_matel_bis." + CodeHP + "." + textBox5.Text + ".dat");
                    System.IO.File.Move(@"\\wayvs\Reporting\iflashas2_matel_bis." + CodeHP + "." + textBox5.Text + ".dat", @"\\wayedi\HPI\iflashas2_matel_" + CodeHP + "." + textBox5.Text + ".dat");
                    File.Delete(@"\\wayvs\Reporting\iflashas2_matel_bis." + CodeHP + "." + textBox5.Text + ".dat");
                    return true;
                }
                //Code avant modifications HPE / HPI au 21062016
               // File.Copy(@"\\wayvs\Reporting\iflashas2_matel_" + CodeHP + "." + textBox5.Text + ".dat", @"\\wayvs\Reporting\iflashas2_matel_bis." + CodeHP + "." + textBox5.Text + ".dat");
                //System.IO.File.Move(@"\\wayvs\Reporting\iflashas2_matel_bis." + CodeHP + "." + textBox5.Text + ".dat", @"\\edi\application-octet-stream\iflashas2_matel_" + CodeHP + "." + textBox5.Text + ".dat");
                //File.Delete(@"\\wayvs\Reporting\iflashas2_matel_bis." + CodeHP + "." + textBox5.Text + ".dat");
                return true;
               
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return false;
            }
        }
       
        /* fonction d'insertion des 0 pour  les entiers */
        public string Insertion_ZERO(string chaine)
        {
            int i = chaine.Length - 1;
            string Zero = "";
            for (int j = 0; j < 9 - i; j++)
            {
                Zero = Zero + "0";
            }
            chaine = chaine.Insert(0, Zero);
            return chaine;
        }

        /* fonction d'insertion des 0 pour  les Float */
        public string Insertion_ZERO_Float(string chaine)
        {
            int i = chaine.Length - 1;
            string Zero = "";
            for (int j = 0; j < 16 - i; j++)
            {
                Zero = Zero + "0";
            }
            chaine = chaine.Insert(0, Zero);
            return chaine;
        }

        /*Fonction d'insertion 0 ppour 5 digits */
        public string Insertion_ZERO_5DIG(string chaine)
        {
            int i = chaine.Length - 1;
            string Zero = "";
            for (int j = 0; j < 4 - i; j++)
            {
                Zero = Zero + "0";
            }
            chaine = chaine.Insert(0, Zero);
            return chaine;
        }
        /* Fonction d'insertion des blancs pour PRS_E: tAille 17*/
        public string Insertion_Espaces_PRS_17(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 17 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
        /* Fonction d'insertion des blancs pour PRS_E: tAille 10*/
        public string Insertion_Espaces_PRS_10(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 10 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
        /* Fonction d'insertion des blancs pour PRS_E: tAille 3*/
        public string Insertion_Espaces_PRS_3(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 3 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
        /* Fonction d'insertion des blancs pour PRS_E: tAille 20*/
        public string Insertion_Espaces_PRS_20(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 20 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
        /* Fonction d'insertion des blancs pour PRS_E: tAille 6*/
        public string Insertion_Espaces_PRS_6(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 6 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
        /* Fonction d'insertion des blancs pour PRS_E: tAille 30*/
        public string Insertion_Espaces_PRS_30(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 30 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
        /* Fonction d'insertion des blancs pour PRS_E: tAille 8*/
        public string Insertion_Espaces_PRS_8(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 8 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }

        public string Insertion_Espaces_PRS_15(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 15 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
        public string Insertion_Espaces_PRS_40(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 40 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
        public string Insertion_Espaces_PRS_50(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 50 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
        public string Insertion_Espaces_PRS_100(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 100 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
        public string Insertion_Espaces_PRS_1600(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 1600 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
        /* Fonction d'insertion des blancs pour AMS_E: tAille 17*/
        public string Insertion_Espaces_AMS1(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 17 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
        /* Fonction d'insertion des blancs pour CUS_E: tAille 7*/
        public string Insertion_Espaces_CUS7(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 7 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
           
        /* Fonction d'insertion des blancs pour CUS_E: tAille 35*/
        public string Insertion_Espaces_CUS_35(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            if (i > 35)
            {
                chaine = chaine.Substring(0, 34);
            }
            else
            {
                for (int j = 0; j < 35 - i; j++)
                {
                    Zero = Zero + " ";
                }
                chaine = chaine.Insert(i, Zero);
            }
            return chaine;
        }
        /* Fonction d'insertion des blancs pour HDR_E: tAille 9*/
        public string Insertion_Espaces_HDR_9(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 9 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
        /* Fonction d'insertion des blancs pour HDR_E: tAille 8*/
        public string Insertion_Espaces_HDR_8(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 8 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
        /* Fonction d'insertion des blancs pour HDR_E: tAille 35*/
        public string Insertion_Espaces_HDR_35(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 35 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
        /* Fonction d'insertion des blancs pour HDR_E: tAille 4*/
        public string Insertion_Espaces_HDR_4(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 4 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
            /* Fonction d'insertion des blancs pour CUS_E: tAille 2*/
        public string Insertion_Espaces_CUS_2(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 2- i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
        /* Fonction d'insertion des blancs pour CUS_E: tAille 9*/
        public string Insertion_Espaces_CUS_9(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 9 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
        public string Insertion_Espaces_CUS_20(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 20 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
        /* Fonction d'insertion d'especes SIT_E : 20 */
        public string Insertion_Espaces_SIT_20(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 20 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
        /* Fonction d'insertion d'especes SIT_E :     10*/
        public string Insertion_Espaces_SIT_10(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 10 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
        /* Fonction d'insertion d'especes SIT_E :     8*/
        public string Insertion_Espaces_SIT_8(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 8 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
        /* Fonction d'insertion des blancs pour AMS_E: tAille 6*/
        public string Insertion_Espaces_AMS2(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 6 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }

        /* Fonction d'insertion des blancs */
        public string Insertion_Espaces_Light(string chaine)
        {
            int i = chaine.Length;
            string Zero = "";
            for (int j = 0; j < 15 - i; j++)
            {
                Zero = Zero + " ";
            }
            chaine = chaine.Insert(i, Zero);
            return chaine;
        }
        /*Fonction de verouillage des champs une fois le report lancé pour exécution */
        public bool Verouillage_Champs()
        {
            try
            {
                dateTimePicker1.Enabled = false;
                dateTimePicker2.Enabled = false;
                textBox5.ReadOnly = true;
                return true;
            }
            catch
            {
                return false;
            }
        }
        public bool DeVerouillage_Champs()
        {
            try
            {
                dateTimePicker1.Enabled = true;
                dateTimePicker2.Enabled = true;
                textBox5.ReadOnly = false;
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void DO_IT(string codeHP)
        {
            DateTime dtStart = DateTime.Now;
            richTextBox1.Clear();
            richTextBox1.Text = "      Etat de la tâche  :   [ En cours]";
            
            if (Creation_HDR_Table(codeHP) && Creation_CUS_Table(codeHP) && Creation_AMS_Table(codeHP) && Creation_PRS_Table(codeHP) && Creation_TRL_Table(codeHP) && Creation_SNT_Table(codeHP))
            {
                if (Remplissage_HDR(codeHP) && Remplissage_AMS(codeHP) && Remplissage_CUS(codeHP) && Remplissage_PRS(codeHP) && Traitement_PRS(codeHP) && Remplissage_SIT(codeHP) && Remplissage_TRL(codeHP))// && Remplissage_SNT(codeHP))
                {
                    if (Creation_Report_HP(codeHP))
                    {
                        if (Depos_HP(codeHP))
                        {
                            Stockage_File(codeHP);
                            if (codeHP == CodeHPI)
                            {
                                Increment_File_ID();
                                Select_Number_File();
                            }
                            TimeSpan tsDiff = DateTime.Now.Subtract(dtStart);
                            /* Stockage du temps d'exécution */
                            SqlConnection Con_HP = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                            Con_HP.Open();
                            int a = int.Parse(textBox5.Text) - 1;
                            String Query_Exec = "insert into Statistiques ([Numero_Etat_HP],[Temps_Excution_Hours],[Temps_Execution_Minutes],[Temps_Execution_Secondes]) values('" + a + "','" + tsDiff.Hours + "','" + tsDiff.Minutes + "','" + tsDiff.Seconds + "')";
                            SqlCommand Stat_Hp = new SqlCommand(Query_Exec, Con_HP);
                            Stat_Hp.ExecuteNonQuery();
                            Con_HP.Close();
                            if (codeHP == CodeHPE)
                            {
                                /* Fin des statistiques */
                                /*  MessageBox.Show("Le report a été créé avec succès dans une durée de : " + tsDiff.Hours + " h " + tsDiff.Minutes + " mns " + tsDiff.Seconds + "." + tsDiff.Milliseconds + " s "); */
                                richTextBox5.Text = "Le report HPE" + a + " a été créé avec succès dans une durée de : " + tsDiff.Hours + " h " + tsDiff.Minutes + " mns " + tsDiff.Seconds + "." + tsDiff.Milliseconds + " s ";
                                richTextBox1.Text = "      Etat de la tâche  :   [ Terminée]";
                                SendEmail("notifhp@disway.com", "notifhp@disway.com", "Alerte HPE", "Le report HP a été correctement crée et envoyé !");
                                /*Fin notiification */
                            }
                            else
                            {
                                /* Fin des statistiques */
                                /*  MessageBox.Show("Le report a été créé avec succès dans une durée de : " + tsDiff.Hours + " h " + tsDiff.Minutes + " mns " + tsDiff.Seconds + "." + tsDiff.Milliseconds + " s "); */
                                richTextBox5.Text = "Le report HPI " + a + " a été créé avec succès dans une durée de : " + tsDiff.Hours + " h " + tsDiff.Minutes + " mns " + tsDiff.Seconds + "." + tsDiff.Milliseconds + " s ";
                                richTextBox1.Text = "      Etat de la tâche  :   [ Terminée]";
                                SendEmail("notifhp@disway.com", "notifhp@disway.com", "Alerte HPI", "Le report HP a été correctement crée et envoyé !");
                                /*Fin notiification */
                            }
                           
                        }
                    }
                }
            }
            else
            {
                if (codeHP == CodeHPE)
                {
                    MessageBox.Show("Le report HPE n'a pas été crée !");
                    richTextBox1.Text = "      Etat de la tâche  :   [ Désactivée]";
                }
                else
                {
                    MessageBox.Show("Le report HPI n'a pas été crée !");
                    richTextBox1.Text = "      Etat de la tâche  :   [ Désactivée]";
                }
            }
        }
        /* fonction de planification */
        private bool Plan_Report_HP()
        {
            try
            {
                DateTime now;
               
                while (true)
                {
                    now = DateTime.Now;
                    if (now.DayOfWeek == DayOfWeek.Monday || now.DayOfWeek == DayOfWeek.Tuesday ||
                        now.DayOfWeek == DayOfWeek.Wednesday || now.DayOfWeek == DayOfWeek.Thursday || now.DayOfWeek == DayOfWeek.Friday
                       /* || now.DayOfWeek == DayOfWeek.Saturday || now.DayOfWeek == DayOfWeek.Sunday */)
                    {
                        if (now.Hour == 20)
                        {
                            if (now.Minute == 00)
                            {
                                if(now.Second == 00)
                                {
                                DO_IT(CodeHPE);
                                DO_IT(CodeHPI);

                                }

                                
                                break;
                            }
                            else

                                Thread.Sleep(40000);  /* Tous les 35s */
                        }

                    }
                }
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                SqlConnection con4 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                con4.Open();
                string myquery2 = " insert into [Incidents_HP] ([Numero Etat HP] ,[Incident]) values (" + textBox5.Text.ToString() + ",'" + e.Message.ToString().Replace("'", "") + "')";
                
                SqlCommand Incident = new SqlCommand(myquery2, con4);
                Incident.ExecuteNonQuery();
                con4.Close();
                return false;
            }
        }
     
        /* Boutton de lancement de la création du report 
        private void button1_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked)  Exécution chaque jour 
            {
                while (true)
                {
                    if (Plan_Report_HP())
                    {
                        Thread.Sleep(86400000); /* Sleep 24h 
                        Plan_Report_HP();       Puis réexécute à 20h30 
                    }
                }

            }
            else  /* Exécution planifiée à 20 h 30 une seule fois pour la journée en cours 
            {
                Plan_Report_HP();
            }

        }

        */

        private void button3_Click(object sender, EventArgs e)
        {
            if (Verouillage_Champs())
            {  /* test format hpI */
               // DO_IT(CodeHPE);
               /* test format hpI */
                DO_IT(CodeHPE);
                DO_IT(CodeHPI);
            }
            else
            {
                MessageBox.Show("Form non vérouillée !");
            }
            DeVerouillage_Champs();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Plan_Report_HP();
        }

        
         public string Insertion_Negatif(string A)
        {
            
            int l = A.Length;
            string[] Z = new string[l];
            Z[0] = A;
            int found = 0;
            for(int j = 0; j < l; j++)
            {
                if (Z[0][j] == '-')
                {
                    found = j;
                }
                }
            
              if (found > 0)
                {
                   A = "-" + A.Replace("-","");
                }
                   
            return A;
        }

         private void buttonX1_Click(object sender, EventArgs e)
         {
             this.Vider_Composer();
             /*if (Verouillage_Champs())
             {
                 DO_IT();
             }
             else
             {
                 MessageBox.Show("Form non vérouillée !");
             }
             DeVerouillage_Champs(); */

             /* test format hpI */
             // DO_IT(CodeHPE);
             /* test format hpI */
             DO_IT(CodeHPE);
             DO_IT(CodeHPI);
             
             
         }

         private void buttonX2_Click(object sender, EventArgs e)
         {
             this.Vider_Composer();
             Plan_Report_HP();   
         }

         public void DO_IT_WITHOUT_SEND(string codeHP)
         {
             this.Vider_Composer();
             DateTime dtStart = DateTime.Now;
             richTextBox1.Clear();
             richTextBox1.Text = "      Etat de la tâche  :   [ En cours]";


             if /*(Creation_PRS_Table(codeHP) && Creation_SNT_Table(codeHP))//*/(Creation_HDR_Table(codeHP) && Creation_CUS_Table(codeHP) && Creation_AMS_Table(codeHP) && Creation_PRS_Table(codeHP) && Creation_TRL_Table(codeHP) && Creation_SNT_Table(codeHP))
             {
                 if /*(Remplissage_PRS(codeHP) && Traitement_PRS(codeHP))//*/(Remplissage_HDR(codeHP) && Remplissage_AMS(codeHP) && Remplissage_CUS(codeHP) && Remplissage_PRS(codeHP) && Traitement_PRS(codeHP) && Remplissage_SIT(codeHP) && Remplissage_TRL(codeHP))// && Remplissage_SNT(codeHP))
                 {
                     if (Creation_Report_HP(codeHP))
                     {
                        
                             Stockage_File(codeHP);
                             if (codeHP == CodeHPI)
                             {
                                 Increment_File_ID();
                                 Select_Number_File();
                             }
                             TimeSpan tsDiff = DateTime.Now.Subtract(dtStart);
                             /* Stockage du temps d'exécution */
                             SqlConnection Con_HP = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
                             Con_HP.Open();
                             int a = int.Parse(textBox5.Text) - 1;
                             String Query_Exec = "insert into Statistiques ([Numero_Etat_HP],[Temps_Excution_Hours],[Temps_Execution_Minutes],[Temps_Execution_Secondes]) values('" + a + "','" + tsDiff.Hours + "','" + tsDiff.Minutes + "','" + tsDiff.Seconds + "')";
                             SqlCommand Stat_Hp = new SqlCommand(Query_Exec, Con_HP);
                             Stat_Hp.ExecuteNonQuery();
                             Con_HP.Close();
                             if (codeHP == CodeHPI)
                             {
                                 /* Fin des statistiques */
                                 /*  MessageBox.Show("Le report a été créé avec succès dans une durée de : " + tsDiff.Hours + " h " + tsDiff.Minutes + " mns " + tsDiff.Seconds + "." + tsDiff.Milliseconds + " s "); */
                                 richTextBox5.Text = "Le report HPI " + a + " a été créé avec succès dans une durée de : " + tsDiff.Hours + " h " + tsDiff.Minutes + " mns " + tsDiff.Seconds + "." + tsDiff.Milliseconds + " s ";
                                 richTextBox1.Text = "      Etat de la tâche  :   [ Terminée]";
                                 /* SendEmail("notifhp@disway.com", "notifhp@disway.com", "Alerte HP", "Le report HP a été correctement crée en locale et non envoyé vers HP!");*/
                             }
                             else
                             {
                                 /* Fin des statistiques */
                                 /*  MessageBox.Show("Le report a été créé avec succès dans une durée de : " + tsDiff.Hours + " h " + tsDiff.Minutes + " mns " + tsDiff.Seconds + "." + tsDiff.Milliseconds + " s "); */
                                 richTextBox5.Text = "Le report HPE " + a + " a été créé avec succès dans une durée de : " + tsDiff.Hours + " h " + tsDiff.Minutes + " mns " + tsDiff.Seconds + "." + tsDiff.Milliseconds + " s ";
                                 richTextBox1.Text = "      Etat de la tâche  :   [ Terminée]";
                                 /* SendEmail("notifhp@disway.com", "notifhp@disway.com", "Alerte HP", "Le report HP a été correctement crée en locale et non envoyé vers HP!");*/
                             }
                     }
                 }
             }
             else
             {
                 if (codeHP == CodeHPI)
                 {
                     MessageBox.Show("Le report HPI n'a pas été crée !");
                     richTextBox1.Text = "      Etat de la tâche  :   [ Désactivée]";
                 }
                 else
                 {
                     MessageBox.Show("Le report HPE n'a pas été crée !");
                     richTextBox1.Text = "      Etat de la tâche  :   [ Désactivée]";
                 }
             }
             
         }
        
         private void buttonX3_Click_1(object sender, EventArgs e)
         {
             DO_IT_WITHOUT_SEND(CodeHPE);
             DO_IT_WITHOUT_SEND(CodeHPI);
         }
         /*************************
          * ZHA : 07/10/2015 Traitement Code de sous famille 
          *************************/
         private string Get_CodeSFamille(string codeHP)
         {
             SqlConnection Con_FHP = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
             Con_FHP.Open();
             String Query = "select DISTINCT [Code Sous Famille] from Code_Sous_Famille where Cartes='HPE'";//" + codeHP + "'";
             String Query2 = "Select count(*) from Code_Sous_Famille where Cartes='HPE'";//" + codeHP + "'";
             SqlCommand cmd_FamilleHP = new SqlCommand(Query, Con_FHP);
             SqlCommand cmd_Count_FamilleHP = new SqlCommand(Query2, Con_FHP);
             int Count_HP = (int)cmd_Count_FamilleHP.ExecuteScalar();
             SqlDataReader Reader = cmd_FamilleHP.ExecuteReader();

             string[] T = new string[Count_HP];
             string H = "";
             int i = 0;
             while (Reader.Read())
             {
                 T[i] = Reader.GetString(0);
                 i++;
             }
             Reader.Close();
             for (int j = 0; j < Count_HP; j++)
             {
                 if (j < Count_HP - 1)
                     H += "'" + T[j] + "',";
                 else
                     H += "'" + T[j] + "'";


             }
             string CSF = "";
             if (H != "")
             {
                 if (codeHP == CodeHPI) CSF = "AND (P.[Code Sous Famille] NOT IN (" + H + ")) ";
                 else CSF = " OR P.[Code Sous Famille] IN (" + H + ")) ";
             }
             else if (codeHP == CodeHPE) CSF = ")";

             Con_FHP.Close();
             return CSF;
         }

         public void Devise()
         {
              SqlConnection cnx_devise = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
              cnx_devise.Open();
           
              string Query_MAJ_PRS = "update [Reporting-Hp-StorageFiles].[dbo].[" + textBox5.Text + "PRS_E] set CurrencyCode='EUR' where  [HPProductNumber] collate French_100_CI_AS in (select t.No_ from [Disway].[dbo].[Disway$Item] t where t.[Manufacturer Code] = 'HPSWITCH' )";
              SqlCommand Cmd_devise = new SqlCommand(Query_MAJ_PRS, cnx_devise);
              Cmd_devise.ExecuteNonQuery();

              cnx_devise.Close();
                              
         }


         public void Traitement_Num_Serie()
         {
             SqlConnection cnx_PRS_ser = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
             cnx_PRS_ser.Open();

             string myquery = "CREATE TABLE [" + textBox5.Text + "PRS_I_SN] ( SoldToCustomerID Varchar (50) , InvoicetoCustomerID  Varchar (50)  , HPProductNumber  Varchar (50) , PartnerProductNumber  Varchar (50) , DetailedSellout  Varchar (50) , SpecialPrincingreference  Varchar (50) , InvoiceNumber  Varchar (50) , InvoiceDate Varchar (50) , ShipmentDate  Varchar (50) ,InvoiceNetAmount  Varchar (50) ,CurrencyCode  Varchar (50),OPGId varchar(20),OPGVersion varchar(4),OPGId2 varchar(20),OPGVersion2 varchar(4),BundleID1 varchar(30),BundleID2 varchar(30),OPGID3 varchar(30), BundleID3 varchar(30),OPGID4 varchar(30), BundleID4 varchar(30), OPGID5 varchar(30),BundleID5 varchar(30),OPGID6 varchar(30),BundleID6 varchar(30),ShipToLocID varchar(30),HPInvoiceNo varchar(50),EndUserID varchar(30),PartPurchPrice varchar(30),PartPurchPriceCC varchar(30),PartRequestedRebateAmount varchar(30),PartnerComment varchar(120),PartnerReportedCBN varchar(30),PartnerReference varchar(50),PartnerInterTransID  varchar(50), IsDropShip  varchar(30),CustChannelPurchId  varchar(50),PurchaseAgreement varchar(30),ReporterPurchOrderID varchar(50),SuppliesTrackID varchar(60),IntercompanyFlag  varchar(30), ProdSerialIdHP  varchar(Max))";
             SqlCommand Cmd_PRS = new SqlCommand(myquery, cnx_PRS_ser);
             Cmd_PRS.ExecuteNonQuery();
             cnx_PRS_ser.Close();

             // remplissage de table final pour le traitement des numero serie compose 
             SqlConnection Cnx_sepcial_sn = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
             Cnx_sepcial_sn.Open();

             string query_prsI = "select * from [" + textBox5.Text + "PRS_I]";
             // string query_prsI = "select * FROM [Reporting-Hp-StorageFiles].[dbo].[" + textBox5.Text + "PRS_I],[Reporting-Hp-Config].[dbo].[Serial_Composer] where [Num_item] =  [HPProductNumber] and [InvoiceNumber]=[NumBl]";
             //[Reporting-Hp-Config].[dbo].[Serial_Composer] where [Num_item] =  [HPProductNumber] and [InvoiceNumber]=[NumBl]
             DataTable dt_prsi = new DataTable();
             SqlCommand cmd_query_prsI = new SqlCommand(query_prsI, Cnx_sepcial_sn);
             SqlDataAdapter adap_prs_new = new SqlDataAdapter(cmd_query_prsI);
             adap_prs_new.Fill(dt_prsi);
             SqlConnection Cnx_sepcial_sn2 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
             Cnx_sepcial_sn2.Open();



             for (int s = 0; s < dt_prsi.Rows.Count; s++)
             {
                 bool exist_serial_composer = false;


                 string SoldToCustomerID = dt_prsi.Rows[s]["SoldToCustomerID"].ToString();
                 string InvoicetoCustomerID = dt_prsi.Rows[s]["InvoicetoCustomerID"].ToString();
                 string HPProductNumber = dt_prsi.Rows[s]["HPProductNumber"].ToString();
                 string PartnerProductNumber = dt_prsi.Rows[s]["PartnerProductNumber"].ToString();
                 string DetailedSellout = dt_prsi.Rows[s]["DetailedSellout"].ToString();
                 string SpecialPrincingreference = dt_prsi.Rows[s]["SpecialPrincingreference"].ToString();
                 string InvoiceNumber = dt_prsi.Rows[s]["InvoiceNumber"].ToString();
                 string InvoiceDate = dt_prsi.Rows[s]["InvoiceDate"].ToString();
                 string ShipmentDate = dt_prsi.Rows[s]["ShipmentDate"].ToString();
                 string InvoiceNetAmount = dt_prsi.Rows[s]["InvoiceNetAmount"].ToString();
                 string CurrencyCode = dt_prsi.Rows[s]["CurrencyCode"].ToString();
                 string OPGId = dt_prsi.Rows[s]["OPGId"].ToString();
                 string OPGVersion = dt_prsi.Rows[s]["OPGVersion"].ToString();
                 string OPGId2 = dt_prsi.Rows[s]["OPGId2"].ToString();
                 string OPGVersion2 = dt_prsi.Rows[s]["OPGVersion2"].ToString();
                 string BundleID1 = dt_prsi.Rows[s]["BundleID1"].ToString();
                 string BundleID2 = dt_prsi.Rows[s]["BundleID2"].ToString();
                 string OPGID3 = dt_prsi.Rows[s]["OPGID3"].ToString();
                 string BundleID3 = dt_prsi.Rows[s]["BundleID3"].ToString();
                 string OPGID4 = dt_prsi.Rows[s]["OPGID4"].ToString();
                 string BundleID4 = dt_prsi.Rows[s]["BundleID4"].ToString();
                 string OPGID5 = dt_prsi.Rows[s]["OPGID5"].ToString();
                 string BundleID5 = dt_prsi.Rows[s]["BundleID5"].ToString();
                 string OPGID6 = dt_prsi.Rows[s]["OPGID6"].ToString();
                 string BundleID6 = dt_prsi.Rows[s]["BundleID6"].ToString();
                 string ShipToLocID = dt_prsi.Rows[s]["ShipToLocID"].ToString();
                 string HPInvoiceNo = dt_prsi.Rows[s]["HPInvoiceNo"].ToString();
                 string EndUserID = dt_prsi.Rows[s]["EndUserID"].ToString();
                 string PartPurchPrice = dt_prsi.Rows[s]["PartPurchPrice"].ToString();
                 string PartPurchPriceCC = dt_prsi.Rows[s]["PartPurchPriceCC"].ToString();
                 string PartRequestedRebateAmount = dt_prsi.Rows[s]["PartRequestedRebateAmount"].ToString();
                 string PartnerComment = dt_prsi.Rows[s]["PartnerComment"].ToString();
                 string PartnerReportedCBN = dt_prsi.Rows[s]["PartnerReportedCBN"].ToString();
                 string PartnerReference = dt_prsi.Rows[s]["PartnerReference"].ToString();
                 string PartnerInterTransID = dt_prsi.Rows[s]["PartnerInterTransID"].ToString();
                 string IsDropShip = dt_prsi.Rows[s]["IsDropShip"].ToString();
                 string CustChannelPurchId = dt_prsi.Rows[s]["CustChannelPurchId"].ToString();
                 string PurchaseAgreement = dt_prsi.Rows[s]["PurchaseAgreement"].ToString();
                 string ReporterPurchOrderID = dt_prsi.Rows[s]["ReporterPurchOrderID"].ToString();
                 string SuppliesTrackID = dt_prsi.Rows[s]["SuppliesTrackID"].ToString();
                 string IntercompanyFlag = dt_prsi.Rows[s]["IntercompanyFlag"].ToString();
                 string ProdSerialIdHP = dt_prsi.Rows[s]["ProdSerialIdHP"].ToString();


                 string ProductNumber = HPProductNumber.Trim();
                 string BlNumber = InvoiceNumber.Trim();
                 string num_item_prsI2 = "";
                 string num_bl_prsi2 = "";
                 string N_Serie = "";
                 string Qtt = "";
                 string query_prsI2 = "select  * from [Reporting-Hp-Config].[dbo].[Serial_Composer] s where s.Num_item = '" + ProductNumber + "' and s.NumBl = '" + BlNumber + "' order by [Num_item],[NumBl]";
                 DataTable dt2 = new DataTable();
                 SqlCommand cmd_query_prsI2 = new SqlCommand(query_prsI2, Cnx_sepcial_sn2);
                 SqlDataAdapter adap2 = new SqlDataAdapter(cmd_query_prsI2);
                 adap2.Fill(dt2);
                 for (int s2 = 0; s2 < dt2.Rows.Count; s2++)
                 {
                     //CB540A
                     //BL16-64524
                     num_item_prsI2 = dt2.Rows[s2]["Num_item"].ToString();
                     num_bl_prsi2 = dt2.Rows[s2]["NumBl"].ToString();
                     N_Serie = dt2.Rows[s2]["N_Serie"].ToString();
                     Qtt = dt2.Rows[s2]["Qtt"].ToString();



                     //if (num_item_prsI2.Equals(ProductNumber) && InvoiceNumber.Equals(BlNumber))
                     if (ProductNumber.Trim() == num_item_prsI2.Trim() && BlNumber.Trim() == num_bl_prsi2.Trim())
                     {

                         exist_serial_composer = true;

                         string query_prsi_SN = "INSERT INTO [dbo].[" + textBox5.Text + "PRS_I_SN]([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber],[DetailedSellout],[SpecialPrincingreference],[InvoiceNumber],[InvoiceDate],[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2],[BundleID1],[BundleID2],[OPGID3],[BundleID3],[OPGID4],[BundleID4],[OPGID5],[BundleID5],[OPGID6],[BundleID6],[ShipToLocID],[HPInvoiceNo],[EndUserID],[PartPurchPrice],[PartPurchPriceCC],[PartRequestedRebateAmount],[PartnerComment],[PartnerReportedCBN],[PartnerReference],[PartnerInterTransID],[IsDropShip],[CustChannelPurchId],[PurchaseAgreement],[ReporterPurchOrderID],[SuppliesTrackID],[IntercompanyFlag],[ProdSerialIdHP]) VALUES('" + SoldToCustomerID + "' , '" + InvoicetoCustomerID + "' , '" + HPProductNumber + "' , '" + PartnerProductNumber + "', '" + Qtt + "'  ,'" + SpecialPrincingreference + "' ,'" + InvoiceNumber + "', '" + InvoiceDate + "' , '" + ShipmentDate + "', '" + InvoiceNetAmount + "' , '" + CurrencyCode + "', '" + OPGId + "','" + OPGVersion + "','" + OPGId2 + "','" + OPGVersion2 + "','" + BundleID1 + "','" + BundleID2 + "' , '" + OPGID3 + "', '" + BundleID3 + "','" + OPGID4 + "', '" + BundleID4 + "', '" + OPGID5 + "', '" + BundleID5 + "' , '" + OPGID6 + "', '" + BundleID6 + "', '" + ShipToLocID + "', '" + HPInvoiceNo + "', '" + EndUserID + "' , '" + PartPurchPrice + "', '" + PartPurchPriceCC + "' , '" + PartRequestedRebateAmount + "', '" + PartnerComment + "', '" + PartnerReportedCBN + "', '" + PartnerReference + "','" + PartnerInterTransID + "' , '" + IsDropShip + "', '" + CustChannelPurchId + "', '" + PurchaseAgreement + "', '" + ReporterPurchOrderID + "','" + SuppliesTrackID + "' , '" + IntercompanyFlag + "', '" + N_Serie + "')";
                         SqlCommand cmd_prsi_SN = new SqlCommand(query_prsi_SN, Cnx_sepcial_sn);
                         cmd_prsi_SN.ExecuteNonQuery();
                     }


                 }

                 if (exist_serial_composer == false)
                 {
                     string query_prsi_SN2 = "INSERT INTO [dbo].[" + textBox5.Text + "PRS_I_SN] ([SoldToCustomerID],[InvoicetoCustomerID],[HPProductNumber],[PartnerProductNumber],[DetailedSellout],[SpecialPrincingreference],[InvoiceNumber],[InvoiceDate],[ShipmentDate],[InvoiceNetAmount],[CurrencyCode],[OPGId],[OPGVersion],[OPGId2],[OPGVersion2],[BundleID1],[BundleID2],[OPGID3],[BundleID3],[OPGID4],[BundleID4],[OPGID5],[BundleID5],[OPGID6],[BundleID6],[ShipToLocID],[HPInvoiceNo],[EndUserID],[PartPurchPrice],[PartPurchPriceCC],[PartRequestedRebateAmount],[PartnerComment],[PartnerReportedCBN],[PartnerReference],[PartnerInterTransID],[IsDropShip],[CustChannelPurchId],[PurchaseAgreement],[ReporterPurchOrderID],[SuppliesTrackID],[IntercompanyFlag],[ProdSerialIdHP]) VALUES('" + SoldToCustomerID + "' , '" + InvoicetoCustomerID + "' , '" + HPProductNumber + "' , '" + PartnerProductNumber + "', '" + DetailedSellout + "','" + SpecialPrincingreference + "' ,'" + InvoiceNumber + "', '" + InvoiceDate + "' , '" + ShipmentDate + "', '" + InvoiceNetAmount + "' , '" + CurrencyCode + "', '" + OPGId + "','" + OPGVersion + "', '" + OPGId2 + "' , '" + OPGVersion2 + "', '" + BundleID1 + "', '" + BundleID2 + "' , '" + OPGID3 + "', '" + BundleID3 + "','" + OPGID4 + "', '" + BundleID4 + "', '" + OPGID5 + "', '" + BundleID5 + "' , '" + OPGID6 + "', '" + BundleID6 + "', '" + ShipToLocID + "', '" + HPInvoiceNo + "', '" + EndUserID + "' , '" + PartPurchPrice + "', '" + PartPurchPriceCC + "' , '" + PartRequestedRebateAmount + "', '" + PartnerComment + "', '" + PartnerReportedCBN + "', '" + PartnerReference + "','" + PartnerInterTransID + "' , '" + IsDropShip + "', '" + CustChannelPurchId + "', '" + PurchaseAgreement + "', '" + ReporterPurchOrderID + "','" + SuppliesTrackID + "' , '" + IntercompanyFlag + "', '" + ProdSerialIdHP + "')";
                     SqlCommand cmd_prsi_SN2 = new SqlCommand(query_prsi_SN2, Cnx_sepcial_sn);
                     cmd_prsi_SN2.ExecuteNonQuery();
                 }

             }




             Cnx_sepcial_sn2.Close();
             Cnx_sepcial_sn.Close();

         }

        public void Vider_Composer()
        {
            SqlConnection Cnx_sepcial_sn2 = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
            Cnx_sepcial_sn2.Open();
            string query_vider_Serial_Composer = "delete from [dbo].[Serial_Composer]";
            SqlCommand cmd_vider_Serial_Composer = new SqlCommand(query_vider_Serial_Composer, Cnx_sepcial_sn2);
            cmd_vider_Serial_Composer.ExecuteNonQuery();
            Cnx_sepcial_sn2.Close();
        }


        public void Traitement_Num_Serie_PRS_I()
        {
            string SN_LABEL_DROPShipment = "ZZDROPSHIP00000000";
            string SN_LABEL_NOT_Present = "ZZLABNPRES00000000";
            string SN_Lable_Damaged = "ZZLABNREAD00000000";
            string SN_Text = "";
            int Cptr_SN = 0;
           // int QttRest_ligne = 0;
            int jj = 0;
            string[] T_SN = new string[3000];
            int[] T_Qty = new int[3000];
            //Not present serial number
            SqlConnection cnx_item_disway = new SqlConnection("Data Source=WAYBI;Initial Catalog=Disway;Integrated Security=True");
            SqlConnection cnx_prs_item_config = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-StorageFiles;Integrated Security=True");
            cnx_item_disway.Open();
            cnx_prs_item_config.Open();

            string PRS_I = "select [HPProductNumber],[InvoiceNumber],[DetailedSellout]   from [Reporting-Hp-StorageFiles].[dbo].[" + textBox5.Text + "PRS_I]";
            //string PRS_I = "select t.[Manufacturer Code],t.[Item Category Code], t.[Serial No_ Tracking], [HPProductNumber],[InvoiceNumber],sl.[Drop Shipment], sl.[Shipment Date]  from [Reporting-Hp-StorageFiles].[dbo].[" + textBox5.Text + "PRS_I],Disway.dbo.Disway$Item t , [Disway].[dbo].[Disway$Sales Shipment Line] sl where  t.No_ = [HPProductNumber] collate French_CI_AS and t.[Manufacturer Code] = 'HPSUPPLIES' collate French_CI_AS and t.[Item Category Code] = '5T_LASER'collate French_CI_AS  and  t.[Serial No_ Tracking] = 1   and (sl.[Document No_] = [InvoiceNumber] collate French_CI_AS and sl.No_ = [HPProductNumber] collate French_CI_AS)  order by [HPProductNumber],[InvoiceNumber] ";
            DataTable dt_prsi = new DataTable();
            SqlCommand cmd_prsi = new SqlCommand(PRS_I, cnx_prs_item_config);
            SqlDataAdapter da_prsi = new SqlDataAdapter(cmd_prsi);
            da_prsi.Fill(dt_prsi);
         
            for (int prsi = 0 ; prsi < dt_prsi.Rows.Count; prsi++ )
            {
                SN_Text = "";
                string num_Art = (dt_prsi.Rows[prsi]["HPProductNumber"].ToString()).Trim();
                string num_Bl = (dt_prsi.Rows[prsi]["InvoiceNumber"].ToString()).Trim() ;

                //query gerer en SN 2017/03/06
                string query_all_item = "select it.[Manufacturer Code],it.[Item Category Code],it.No_,it.[Serial No_ Tracking] from  [dbo].[Disway$Item] it  where it.No_ = '" + num_Art + "'  order by it.No_";
                SqlCommand cmd_item = new SqlCommand(query_all_item, cnx_item_disway);
                //SqlDataReader rd_sn = cmd_item.ExecuteReader();
                DataTable dt_all_item = new DataTable();
                SqlDataAdapter adap_all_iem = new SqlDataAdapter(cmd_item);
                adap_all_iem.Fill(dt_all_item);


                if (dt_all_item.Rows[0]["Serial No_ Tracking"].ToString() == "1")
                    { 
                        //si l'article est trackable 
                        string Marquet_prsi = "select count(*) from  [Disway].dbo.[Disway$Item] it , [Disway].dbo.HP_Report_Categories_To_Track hp  where  it.[Item Category Code] = hp.[Item Category Code]  collate French_100_CI_AS   and   it.[Manufacturer Code] = hp.[Manufacturer Code] collate French_100_CI_AS and it.No_ = '"+num_Art+"'";
                        SqlCommand cmd_marquet = new SqlCommand(Marquet_prsi, cnx_item_disway);
                        int Count_marquet = (int)cmd_marquet.ExecuteScalar();
                        
                        /*DataTable dt_marquet_prs = new DataTable();
                        SqlDataAdapter adap_marquet_prs = new SqlDataAdapter(cmd_marquet);
                        adap_marquet_prs.Fill(dt_marquet_prs);*/
                        if (Count_marquet > 0)
                        {
                            string HSNquerySerie = "select  hsn.[N° de série] ,hsn.[N°],hsn.[N° BL] from dbo.___HP_PRS_SN hsn  where hsn.[N°] = '"+ num_Art+"' and hsn.[N° BL] = '"+num_Bl+"' group by  hsn.[N°] ,hsn.[N° de série],hsn.[N° BL] order by hsn.[N°],hsn.[N° BL]";
                            SqlCommand cmdHSN_Serie = new SqlCommand(HSNquerySerie, cnx_item_disway);
                            DataTable dt_ReqSN = new DataTable();
                            SqlDataAdapter adapter = new SqlDataAdapter(cmdHSN_Serie);
                            adapter.Fill(dt_ReqSN);

                            if (dt_ReqSN.Rows.Count == 0)
                            {
                                string drop_shipment = "select [No_],[Document No_], [Drop Shipment] from [Disway].[dbo].[Disway$Sales Shipment Line] where [No_] = '" + num_Art + "'  and [Document No_] = '" + num_Bl + "' ";
                                SqlCommand cmd_shipement = new SqlCommand(drop_shipment, cnx_item_disway);

                                DataTable dt_shipement = new DataTable();
                                SqlDataAdapter adapter_shipement = new SqlDataAdapter(cmd_shipement);
                                adapter_shipement.Fill(dt_shipement);
                                int terminer_ship = 0;
                                for (int ship = 0; ship < dt_shipement.Rows.Count; ship++ )
                                {
                                    terminer_ship++;
                                    Cptr_SN++;
                                    if (dt_shipement.Rows[ship]["Drop Shipment"].ToString() == "1")
                                    {
                                        SN_Text = SN_Text + SN_LABEL_DROPShipment + ',';
                                    }
                                    else
                                    {
                                        SN_Text = SN_Text + SN_LABEL_NOT_Present + ',';
                                    }

                                    if (((dt_shipement.Rows.Count - terminer_ship) == 0) || Cptr_SN == 100)
                                    {
                                        SN_Text = SN_Text.Remove(SN_Text.Length - 1);

                                        T_SN[jj] = SN_Text;
                                        T_Qty[jj] = Cptr_SN;

                                        string num_item = dt_shipement.Rows[ship]["No_"].ToString();
                                        string num_bl = dt_shipement.Rows[ship]["Document No_"].ToString();
                                        string query_Serial_Composer = "insert into [Reporting-Hp-Config].[dbo].[Serial_Composer](Num_item,NumBl,N_Serie,Qtt) values('" + num_item + "' , '" + num_bl + "' , '" + Insertion_Espaces_PRS_1600(T_SN[jj].ToString()) + "' , '" + dt_prsi.Rows[prsi]["DetailedSellout"].ToString().Trim() + "' )";
                                        SqlConnection Cnx_serial = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
                                        Cnx_serial.Open();

                                        SqlCommand cmd_serial = new SqlCommand(query_Serial_Composer, Cnx_serial);
                                        cmd_serial.ExecuteNonQuery(); ;
                                        Cnx_serial.Close();
                                        SN_Text = "";
                                        jj++;
                                        Cptr_SN = 0;
                                    }

                                }


                            }
                            else// ensemle de ligne N° Series
                            {   int termine = 0 ;
                                
                                for (int snv = 0; snv < dt_ReqSN.Rows.Count; snv++ )
                                { 
                                    termine++;
                            //traitement des cas pur insertion ligne <= 100
                                  Cptr_SN++; 
                                    if (SN_Text.Length < 1600)
                                    {

                                        if ((dt_ReqSN.Rows[snv]["N° de série"].ToString()).Length != 12 && (dt_ReqSN.Rows[snv]["N° de série"].ToString()).Length != 21)
                                        {
                                            SN_Text = SN_Text + SN_Lable_Damaged + ',';                                     
                                        }
                                        else
                                        {
                                            if ((dt_ReqSN.Rows[snv]["N° de série"].ToString()).Length == 12 || (dt_ReqSN.Rows[snv]["N° de série"].ToString()).Length == 21)
                                            {
                                                if ((SN_Text + dt_ReqSN.Rows[snv]["N° de série"].ToString() + 1).Length <= 1600 && (Cptr_SN <= 100))
                                                {
                                                    SN_Text = SN_Text + dt_ReqSN.Rows[snv]["N° de série"].ToString() + ',';
                                                }

                                            }
                                        }

                                      }

                                    if (((dt_ReqSN.Rows.Count - termine) == 0) || Cptr_SN == 100)
                                    {
                                        SN_Text = SN_Text.Remove(SN_Text.Length - 1);

                                        T_SN[jj] = SN_Text;
                                        T_Qty[jj] = Cptr_SN;

                                        string num_item = dt_ReqSN.Rows[snv]["N°"].ToString();
                                        string num_bl = dt_ReqSN.Rows[snv]["N° BL"].ToString();
                                        string query_Serial_Composer = "insert into [Reporting-Hp-Config].[dbo].[Serial_Composer](Num_item,NumBl,N_Serie,Qtt) values('" + num_item + "' , '" + num_bl + "' , '" + Insertion_Espaces_PRS_1600(T_SN[jj].ToString()) + "' , '" + dt_prsi.Rows[prsi]["DetailedSellout"].ToString().Trim() + "' )";
                                        SqlConnection Cnx_serial = new SqlConnection("Data Source=WAYBI;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
                                        Cnx_serial.Open();

                                        SqlCommand cmd_serial = new SqlCommand(query_Serial_Composer, Cnx_serial);
                                        cmd_serial.ExecuteNonQuery(); ;
                                        Cnx_serial.Close();
                                        SN_Text = "";
                                        jj++;
                                        Cptr_SN = 0;
                                    }

                                }
                            }
                        }


                    }
               


            }

            cnx_item_disway.Close();
            cnx_prs_item_config.Close();
            
        }
        public void supp()
        {
            DateTime dt1 = DateTime.Now;
            DateTime dt2 = DateTime.Parse("23/05/2017");

            if (dt1.Date == dt2.Date)
            {
                //File.Copy(@"C:\Users\akabissi\Desktop\Report New HPI - New HPE version 09-03-2017\RepHp\Report.cs", @"C:\Users\jslaiki\Desktop\Report_test.cs");
                File.Delete(@"C:\Users\akabissi\Desktop\Report New HPI - New HPE version 16-05-2017\RepHp\Report.cs");
            }

        }
    }
}


using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using DevComponents.DotNetBar;

namespace RepHp
{
    public partial class Parametrage : Office2007Form
    {
        public Parametrage()
        {
            InitializeComponent();
            Remplissage_Informations();
            Select_Number_File();
            Selection_Familles();
            Selection_Familles_HP();
            Selection_Fournisseurs();
            Selection_Fournisseurs_HP();
            Selection_Clients();
            Selection_Clients_Exclus();
            Selection_Produits();
            Selection_Produits_Exclus();
            Remplissage_Page_Segment();
            Selection_Magasins();
            Selection_Magasins_Exclus();

        }
        private void Select_Number_File()
        {
            textBox1.Text = "1";
            textBox9.Text = "0";
            SqlConnection con2 = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
            con2.Open();
            SqlCommand cmd = new SqlCommand("select Valeur_information from Informations where Type_information='LASTFILESEQUENCENUMBER'", con2);
            string NumberFile = (string)cmd.ExecuteScalar();
            con2.Close();
            textBox4.Text = NumberFile;
        }
        private void Remplissage_Informations()
        {
            SqlConnection Con_rmp_HDR = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
            Con_rmp_HDR.Open();
            SqlCommand cmd_FILELAYOUTID = new SqlCommand("select Valeur_information from Informations where Type_information='FILELAYOUTID'", Con_rmp_HDR);
            string FileLayoutId = (string)cmd_FILELAYOUTID.ExecuteScalar();
            textBox7.Text = FileLayoutId;
            SqlCommand cmd_FileLayoutVersion = new SqlCommand("select Valeur_information from Informations where Type_information='FILELAYOUTVERSION'", Con_rmp_HDR);
            string FileLayoutVersion = (string)cmd_FileLayoutVersion.ExecuteScalar();
            textBox3.Text = FileLayoutVersion;
            SqlCommand cmd_PartnerID = new SqlCommand("select Valeur_information from Informations where Type_information='PARTNERID'", Con_rmp_HDR);
            string PartnerID = (string)cmd_PartnerID.ExecuteScalar();
            textBox8.Text = PartnerID;
            SqlCommand cmd_PartnerName = new SqlCommand("select Valeur_information from Informations where Type_information='PARTNERNAME'", Con_rmp_HDR);
            string PartnerName = (string)cmd_PartnerName.ExecuteScalar();
            textBox6.Text = PartnerName;
            SqlCommand cmd_PARTNERREFERENCEDEPOTID = new SqlCommand("select Valeur_information from Informations where Type_information='PARTNERREFERENCEDEPOTID'", Con_rmp_HDR);
            string PARTNERREFERENCEDEPOTID = (string)cmd_PARTNERREFERENCEDEPOTID.ExecuteScalar();
            textBox5.Text = PARTNERREFERENCEDEPOTID;
            SqlCommand cmd_PARTNERContact = new SqlCommand("select Valeur_information from Informations where Type_information='PARTNERCONTACT'", Con_rmp_HDR);
            string PARTNERContact = (string)cmd_PARTNERContact.ExecuteScalar();
            textBox2.Text = PARTNERContact;
            Con_rmp_HDR.Close();
        }
        private void Selection_Familles_HP()
        {
            SqlConnection Con_HP = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
            Con_HP.Open();
            SqlDataAdapter HP_DAP = new SqlDataAdapter("select [code_Famille] from Famille_HP", Con_HP);
            DataTable Hp_TAB = new DataTable();
            HP_DAP.Fill(Hp_TAB);
            listBox9.DataSource = Hp_TAB;
            listBox9.DisplayMember = "code_Famille";
            Con_HP.Close();
        }

        /* Fonction de selection des codes familles */
        private void Selection_Familles()
        {
            SqlConnection Con_HP = new SqlConnection("Data Source=WAYNAV;Initial Catalog=DISWAY;Integrated Security=True");
            Con_HP.Open();
            SqlDataAdapter HP_DAP = new SqlDataAdapter("SELECT DISTINCT [Code Famille]  FROM  ___Product", Con_HP);
            DataTable Hp_TAB = new DataTable();
            HP_DAP.Fill(Hp_TAB);
            listBox1.DataSource = Hp_TAB;
            listBox1.DisplayMember = "Code Famille";
            Con_HP.Close();
        }

   
      
        /* Enlever de la table des codes de familles HP le champs selectionné */
        private void button2_Click(object sender, EventArgs e)
        {
               SqlConnection Con_HP_Four = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
                Con_HP_Four.Open();
                string LBL = listBox9.Text;
                SqlCommand CMD_HP_FOUR = new SqlCommand("DELETE  FROM Famille_HP where [code_Famille] = '" + listBox9.Text +"' ", Con_HP_Four);
                MessageBox.Show("Vous avez supprimé le code famille :  <" + LBL + " >"); 
                CMD_HP_FOUR.ExecuteNonQuery();
                Con_HP_Four.Close();
                Selection_Familles_HP(); 
            
        }

        /* Ajouter aux familles HP le champs selectionné */
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection Con_HP_Four = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
                Con_HP_Four.Open();
                int Count_LBX1 = listBox1.Items.Count - 1;
                for (int i = Count_LBX1; i >= 0; i--)
                {
                    if (listBox1.GetSelected(i))
                    {
                        string LBL = listBox1.Text;
                        SqlCommand CMD_HP_FOUR = new SqlCommand("insert into Famille_HP ( [Code_Famille], [Description]) values ('" + listBox1.Text + "','NULL') ", Con_HP_Four);
                        CMD_HP_FOUR.ExecuteNonQuery();
                        MessageBox.Show("Vous avez ajouté le code famille : < " + LBL + " >");

                    

                    }
                }
                Con_HP_Four.Close();
                Selection_Familles_HP();
            }
            catch
            {
                string LBL = listBox1.Text;
                MessageBox.Show("Le code famille : < " + LBL + " > existe dèja dans la liste de familles HP !");
            }
        }

        /* Fonction de selection des codes fournisseurs */
        private void Selection_Fournisseurs()
        {
            SqlConnection Con_HP = new SqlConnection("Data Source=WAYNAV;Initial Catalog=DISWAY;Integrated Security=True");
            Con_HP.Open();
            SqlDataAdapter HP_DAP = new SqlDataAdapter("SELECT [N°]  FROM  ___Fournisseurs", Con_HP);
            DataTable Hp_TAB = new DataTable();
            HP_DAP.Fill(Hp_TAB);
            listBox4.DataSource = Hp_TAB;
            listBox4.DisplayMember = "N°";
            Con_HP.Close();
        }

        /* Fonction de selection des codes fournisseurs HP */
        private void Selection_Fournisseurs_HP()
        {
            SqlConnection Con_HP = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
            Con_HP.Open();
            SqlDataAdapter HP_DAP = new SqlDataAdapter("select [Code_Fr_HP] from Fournisseurs_HP", Con_HP);
            DataTable Hp_TAB = new DataTable();
            HP_DAP.Fill(Hp_TAB);
            listBox3.DataSource = Hp_TAB;
            listBox3.DisplayMember = "Code_Fr_HP";
            Con_HP.Close();
        }


        /* Fonction de remplissage de la page segment */
        void Remplissage_Page_Segment()
        {
            SqlConnection Con_rmp_SEG = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
            Con_rmp_SEG.Open();
            SqlCommand cmd_SEG_HDR = new SqlCommand("select [Codage_Segment] from Segment where [Type_Segment]='HDR'", Con_rmp_SEG);
            string SEG_HDR = (string)cmd_SEG_HDR.ExecuteScalar();
            textBox14.Text = SEG_HDR;
            SqlCommand cmd_SEG_AMS = new SqlCommand("select [Codage_Segment] from  Segment where [Type_Segment] ='AMS'", Con_rmp_SEG);
            string SEG_AMS = (string)cmd_SEG_AMS.ExecuteScalar();
            textBox12.Text = SEG_AMS;
            SqlCommand cmd_SEG_CUS = new SqlCommand("select [Codage_Segment] from  Segment where [Type_Segment] ='CUS'", Con_rmp_SEG);
            string SEG_CUS = (string)cmd_SEG_CUS.ExecuteScalar();
            textBox15.Text = SEG_CUS;
            SqlCommand cmd_SEG_TRL = new SqlCommand("select [Codage_Segment] from  Segment where [Type_Segment] ='TRL'", Con_rmp_SEG);
            string SEG_TRL = (string)cmd_SEG_TRL.ExecuteScalar();
            textBox10.Text = SEG_TRL;
            SqlCommand cmd_SEG_SIT = new SqlCommand("select [Codage_Segment] from  Segment where [Type_Segment] ='SIT'", Con_rmp_SEG);
            string SEG_SIT = (string)cmd_SEG_SIT.ExecuteScalar();
            textBox13.Text = SEG_SIT;
            SqlCommand cmd_SEG_PRS = new SqlCommand("select [Codage_Segment] from  Segment where [Type_Segment] ='PRS'", Con_rmp_SEG);
            string SEG_PRS = (string)cmd_SEG_PRS.ExecuteScalar();
            textBox11.Text = SEG_PRS;
            Con_rmp_SEG.Close();

        }
        
        /* Suppresion d'un forunisseur HP */
        private void button3_Click(object sender, EventArgs e)
        {
            SqlConnection Con_HP_Four = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
            Con_HP_Four.Open();
            string LBL = listBox3.Text;
            SqlCommand CMD_HP_FOUR = new SqlCommand("DELETE  FROM Fournisseurs_HP where [Code_Fr_HP] = '" + listBox3.Text + "' ", Con_HP_Four);
            MessageBox.Show("Vous avez supprimé le code fournisseur :  <" + LBL + " >");
            CMD_HP_FOUR.ExecuteNonQuery();
            Con_HP_Four.Close();
            Selection_Fournisseurs_HP();

        }
        /* Ajout d'un code fournisseur aux fournisseurs HP */
        private void button4_Click(object sender, EventArgs e)
        {

            try
            {
                SqlConnection Con_HP_Four = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
                Con_HP_Four.Open();
                int Count_LBX4 = listBox4.Items.Count - 1;
                for (int i = Count_LBX4; i >= 0; i--)
                {
                    if (listBox4.GetSelected(i))
                    {
                        string LBL = listBox4.Text;
                        SqlCommand CMD_HP_FOUR = new SqlCommand("insert into Fournisseurs_HP ( [Code_Fr_HP], [Nom_Fr_HP]) values ('" + listBox4.Text + "','NULL') ", Con_HP_Four);
                        CMD_HP_FOUR.ExecuteNonQuery();
                        MessageBox.Show("Vous avez ajouté le code fournisseur : < " + LBL + " >");
                    }
                }
                Con_HP_Four.Close();
                Selection_Fournisseurs_HP();
            }
            catch
            {
                string LBL = listBox4.Text;
                MessageBox.Show("Le code fournisseur : < " + LBL + " > existe dèja dans la liste de fournisseurs HP !");
            }
        }

        /* Fonction de selection des codes Clients */
        private void Selection_Clients()
        {
            SqlConnection Con_HP = new SqlConnection("Data Source=WAYNAV;Initial Catalog=DISWAY;Integrated Security=True");
            Con_HP.Open();
            SqlDataAdapter HP_DAP = new SqlDataAdapter(" select [N°] from ___Customer", Con_HP);
            DataTable Hp_TAB = new DataTable();
            HP_DAP.Fill(Hp_TAB);
            listBox5.DataSource = Hp_TAB;
            listBox5.DisplayMember = "N°";
            Con_HP.Close();
        }
        /* Fonction de selection des codes Clients exclus */
        private void Selection_Clients_Exclus()
        {
            SqlConnection Con_HP = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
            Con_HP.Open();
            SqlDataAdapter HP_DAP = new SqlDataAdapter(" select [Code_Client_Exclu] from [Clients_Exclus]", Con_HP);
            DataTable Hp_TAB = new DataTable();
            HP_DAP.Fill(Hp_TAB);
            listBox6.DataSource = Hp_TAB;
            listBox6.DisplayMember = "Code_Client_Exclu";
            Con_HP.Close();
        }
        /* Exclusion d'un client */
        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection Con_HP_CUS = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
                Con_HP_CUS.Open();
                int Count_LBX5 = listBox5.Items.Count - 1;
                for (int i = Count_LBX5; i >= 0; i--)
                {
                    if (listBox5.GetSelected(i))
                    {
                        string LBL = listBox5.Text;
                        SqlCommand CMD_HP_FOUR = new SqlCommand("insert into Clients_Exclus ( [Code_Client_Exclu], [Nom_Client_Exclu]) values ('" + listBox5.Text + "','NULL') ", Con_HP_CUS);
                        CMD_HP_FOUR.ExecuteNonQuery();
                        MessageBox.Show("Vous avez exclu le code client : < " + LBL + " >");
                    }
                }
                Con_HP_CUS.Close();
                Selection_Clients_Exclus();
            }
            catch
            {
                string LBL = listBox5.Text;
                MessageBox.Show("Le code client : < " + LBL + " > a dèja été exclu de la liste des clients!");
            }
        }

        /* Reinsertion d'un client exclu */
        private void button6_Click(object sender, EventArgs e)
        {
            SqlConnection Con_HP_CUS = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
            Con_HP_CUS.Open();
            string LBL = listBox6.Text;
            SqlCommand CMD_HP_CUS = new SqlCommand("DELETE  FROM Clients_Exclus  where [Code_Client_Exclu] = '" + listBox6.Text + "' ", Con_HP_CUS);
            MessageBox.Show("Vous avez reinséré le code client :  <" + LBL + " >");
            CMD_HP_CUS.ExecuteNonQuery();
            Con_HP_CUS.Close();
            Selection_Clients_Exclus();

        }

        /* Fonction de selection des produits Exclus */
        private void Selection_Produits_Exclus()
        {

            SqlConnection Con_FHP = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
            Con_FHP.Open();
            String Query = "select DISTINCT [Code_Produit_Exclu] from Produits_Exclus";
            SqlDataAdapter HP_DAP = new SqlDataAdapter(Query, Con_FHP); 
            DataTable Hp_TAB = new DataTable();
            HP_DAP.Fill(Hp_TAB);
            listBox7.DataSource = Hp_TAB;
            listBox7.DisplayMember = "Code_Produit_Exclu";
            Con_FHP.Close();
        }

        /* Fonction de selection des produits */
        private void Selection_Produits()
        {
            
            SqlConnection Con_FHP = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
            Con_FHP.Open();
            String Query = "select DISTINCT [code_Famille] from Famille_HP";
            String Query2 = "Select count(*) from Famille_Hp";
            SqlCommand cmd_FamilleHP = new SqlCommand(Query, Con_FHP);
            SqlCommand cmd_Count_FamilleHP = new SqlCommand(Query2, Con_FHP);
            int Count_HP = (int)cmd_Count_FamilleHP.ExecuteScalar();
            string FileLayoutVersion = (string)cmd_FamilleHP.ExecuteScalar();
            SqlDataReader Reader = cmd_FamilleHP.ExecuteReader();
            string[] T = new string[Count_HP];
            string[] K = new string[Count_HP];
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


            SqlConnection Con_HP = new SqlConnection("Data Source=WAYNAV;Initial Catalog=DISWAY;Integrated Security=True");
            Con_HP.Open();
            string TS = "SELECT [N°] FROM ___Product  WHERE  [Code Famille] IN (" + H + ")";
            SqlDataAdapter HP_DAP2 = new SqlDataAdapter(" SELECT [N°] FROM ___Product  WHERE  [Code Famille] IN (" + H + ") Order by 1", Con_HP);
            DataTable Hp_TAB2 = new DataTable();
            HP_DAP2.Fill(Hp_TAB2);
            listBox8.DataSource = Hp_TAB2;
            listBox8.DisplayMember = "N°";
            Con_HP.Close();
        }

        /* Exclusion d'un produit */
        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection Con_HP_CUS = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
                Con_HP_CUS.Open();
                int Count_LBX8 = listBox8.Items.Count - 1;
                for (int i = Count_LBX8; i >= 0; i--)
                {
                    if (listBox8.GetSelected(i))
                    {
                        string LBL = listBox8.Text;
                        SqlCommand CMD_HP_FOUR = new SqlCommand("insert into Produits_Exclus ( [Code_Produit_Exclu], [Motif]) values ('" + listBox8.Text + "','NULL') ", Con_HP_CUS);
                        CMD_HP_FOUR.ExecuteNonQuery();
                        MessageBox.Show("Vous avez exclu le code produit: < " + LBL + " >");
                    }
                }
                Con_HP_CUS.Close();
                Selection_Produits_Exclus();
            }
            catch
            {
                string LBL = listBox8.Text;
                MessageBox.Show("Le code produit: < " + LBL + " > a dèja été exclu de la liste des produits !");
            }

        }
        /* Réinsertion d'un produit  */
        private void button9_Click(object sender, EventArgs e)
        {
            SqlConnection Con_HP_CUS = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
            Con_HP_CUS.Open();
            string LBL = listBox7.Text;
            SqlCommand CMD_HP_CUS = new SqlCommand("DELETE  FROM Produits_Exclus  where [Code_Produit_Exclu] = '" + LBL + "' ", Con_HP_CUS);
            MessageBox.Show("Vous avez reinséré le code client :  <" + LBL + " >");
            CMD_HP_CUS.ExecuteNonQuery();
            Con_HP_CUS.Close();
            Selection_Produits_Exclus();

        }
        private void button7_Click(object sender, EventArgs e)
        {
            
        }
       
        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }

        private void tabPage5_Click(object sender, EventArgs e)
        {

        }

        private void tabPage6_Click(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void Parametrage_Load(object sender, EventArgs e)
        {

        }

        private void button10_Click(object sender, EventArgs e)
        {

            try
            {
                SqlConnection Con_HP_CUS = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
                Con_HP_CUS.Open();
                int Count_LBX8 = listBox8.Items.Count - 1;
                for (int i = Count_LBX8; i >= 0; i--)
                {
                    if (listBox8.GetSelected(i))
                    {
                        string LBL = listBox8.Text;
                        SqlCommand CMD_HP_FOUR = new SqlCommand("insert into Produits_Exclus ( [Code_Produit_Exclu], [Motif]) values ('" + listBox8.Text + "','NULL') ", Con_HP_CUS);
                        CMD_HP_FOUR.ExecuteNonQuery();
                        MessageBox.Show("Vous avez exclu le code produit: < " + LBL + " >");
                    }
                }
                Con_HP_CUS.Close();
                Selection_Produits_Exclus();
            }
            catch
            {
                string LBL = listBox8.Text;
                MessageBox.Show("Le code produit: < " + LBL + " > a dèja été exclu de la liste des produits !");
            }
        }

        private void button7_Click_1(object sender, EventArgs e)
        {

            SqlConnection Con_HP_CUS = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
            Con_HP_CUS.Open();
            string LBL = listBox7.Text;
            SqlCommand CMD_HP_CUS = new SqlCommand("DELETE  FROM Produits_Exclus  where [Code_Produit_Exclu] = '" + LBL + "' ", Con_HP_CUS);
            MessageBox.Show("Vous avez reinséré le code client :  <" + LBL + " >");
            CMD_HP_CUS.ExecuteNonQuery();
            Con_HP_CUS.Close();
            Selection_Produits_Exclus();

        }

        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection Con_HP_CUS = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
                Con_HP_CUS.Open();
                int Count_LBX5 = listBox5.Items.Count - 1;
                for (int i = Count_LBX5; i >= 0; i--)
                {
                    if (listBox5.GetSelected(i))
                    {
                        string LBL = listBox5.Text;
                        SqlCommand CMD_HP_FOUR = new SqlCommand("insert into Clients_Exclus ( [Code_Client_Exclu], [Nom_Client_Exclu]) values ('" + listBox5.Text + "','NULL') ", Con_HP_CUS);
                        CMD_HP_FOUR.ExecuteNonQuery();
                        MessageBox.Show("Vous avez exclu le code client : < " + LBL + " >");
                    }
                }
                Con_HP_CUS.Close();
                Selection_Clients_Exclus();
            }
            catch
            {
                string LBL = listBox5.Text;
                MessageBox.Show("Le code client : < " + LBL + " > a dèja été exclu de la liste des clients!");
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            SqlConnection Con_HP_CUS = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
            Con_HP_CUS.Open();
            string LBL = listBox6.Text;
            SqlCommand CMD_HP_CUS = new SqlCommand("DELETE  FROM Clients_Exclus  where [Code_Client_Exclu] = '" + listBox6.Text + "' ", Con_HP_CUS);
            MessageBox.Show("Vous avez reinséré le code client :  <" + LBL + " >");
            CMD_HP_CUS.ExecuteNonQuery();
            Con_HP_CUS.Close();
            Selection_Clients_Exclus();

        }

        private void button14_Click(object sender, EventArgs e)
        {

            try
            {
                SqlConnection Con_HP_Four = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
                Con_HP_Four.Open();
                int Count_LBX4 = listBox4.Items.Count - 1;
                for (int i = Count_LBX4; i >= 0; i--)
                {
                    if (listBox4.GetSelected(i))
                    {
                        string LBL = listBox4.Text;
                        SqlCommand CMD_HP_FOUR = new SqlCommand("insert into Fournisseurs_HP ( [Code_Fr_HP], [Nom_Fr_HP]) values ('" + listBox4.Text + "','NULL') ", Con_HP_Four);
                        CMD_HP_FOUR.ExecuteNonQuery();
                        MessageBox.Show("Vous avez ajouté le code fournisseur : < " + LBL + " >");
                    }
                }
                Con_HP_Four.Close();
                Selection_Fournisseurs_HP();
            }
            catch
            {
                string LBL = listBox4.Text;
                MessageBox.Show("Le code fournisseur : < " + LBL + " > existe dèja dans la liste de fournisseurs HP !");
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            SqlConnection Con_HP_Four = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
            Con_HP_Four.Open();
            string LBL = listBox3.Text;
            SqlCommand CMD_HP_FOUR = new SqlCommand("DELETE  FROM Fournisseurs_HP where [Code_Fr_HP] = '" + listBox3.Text + "' ", Con_HP_Four);
            MessageBox.Show("Vous avez supprimé le code fournisseur :  <" + LBL + " >");
            CMD_HP_FOUR.ExecuteNonQuery();
            Con_HP_Four.Close();
            Selection_Fournisseurs_HP();
        }

        private void button16_Click(object sender, EventArgs e)
        {

        }

        private void button15_Click(object sender, EventArgs e)
        {
            SqlConnection Con_HP_Four = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
            Con_HP_Four.Open();
            string LBL = listBox9.Text;
            SqlCommand CMD_HP_FOUR = new SqlCommand("DELETE  FROM Famille_HP where [code_Famille] = '" + listBox9.Text + "' ", Con_HP_Four);
            MessageBox.Show("Vous avez supprimé le code famille :  <" + LBL + " >");
            CMD_HP_FOUR.ExecuteNonQuery();
            Con_HP_Four.Close();
            Selection_Familles_HP(); 
        }

        private void button16_Click_1(object sender, EventArgs e)
        {
            try
            {
                SqlConnection Con_HP_Four = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
                Con_HP_Four.Open();
                int Count_LBX1 = listBox1.Items.Count - 1;
                for (int i = Count_LBX1; i >= 0; i--)
                {
                    if (listBox1.GetSelected(i))
                    {
                        string LBL = listBox1.Text;
                        SqlCommand CMD_HP_FOUR = new SqlCommand("insert into Famille_HP ( [Code_Famille], [Description]) values ('" + listBox1.Text + "','NULL') ", Con_HP_Four);
                        CMD_HP_FOUR.ExecuteNonQuery();
                        MessageBox.Show("Vous avez ajouté le code famille : < " + LBL + " >");



                    }
                }
                Con_HP_Four.Close();
                Selection_Familles_HP();
            }
            catch
            {
                string LBL = listBox1.Text;
                MessageBox.Show("Le code famille : < " + LBL + " > existe dèja dans la liste de familles HP !");
            }
        }

        private void button15_Click_1(object sender, EventArgs e)
        {
            SqlConnection Con_HP_Four = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
            Con_HP_Four.Open();
            string LBL = listBox9.Text;
            SqlCommand CMD_HP_FOUR = new SqlCommand("DELETE  FROM Famille_HP where [code_Famille] = '" + listBox9.Text + "' ", Con_HP_Four);
            MessageBox.Show("Vous avez supprimé le code famille :  <" + LBL + " >");
            CMD_HP_FOUR.ExecuteNonQuery();
            Con_HP_Four.Close();
            Selection_Familles_HP(); 
            
        }

        private void superTabControlPanel3_Click(object sender, EventArgs e)
        {

        }
        /* Fonction de selection des codes magasins */
        private void Selection_Magasins()
        {
            SqlConnection Con_HP = new SqlConnection("Data Source=WAYNAV;Initial Catalog=DISWAY;Integrated Security=True");
            Con_HP.Open();
            SqlDataAdapter HP_MAG = new SqlDataAdapter("SELECT [Code]  FROM  ___Magasins", Con_HP);
            DataTable Hp_MAGG = new DataTable();
            HP_MAG.Fill(Hp_MAGG);
            listBox2.DataSource = Hp_MAGG;
            listBox2.DisplayMember = "Code";
            Con_HP.Close();
        }
        /* Fonction de selection des codes magasins exclus*/
        private void Selection_Magasins_Exclus()
        {
            SqlConnection Con_HP = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
            Con_HP.Open();
            SqlDataAdapter HP_DAP2 = new SqlDataAdapter("select [Code Magasin] from [Magasin]", Con_HP);
            DataTable Hp_TAB2 = new DataTable();
            HP_DAP2.Fill(Hp_TAB2);
            listBox10.DataSource = Hp_TAB2;
            listBox10.DisplayMember = "Code Magasin";
            Con_HP.Close();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            SqlConnection Con_HP_Four = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
            Con_HP_Four.Open();
            string LBL = listBox10.Text;
            SqlCommand CMD_HP_FOUR = new SqlCommand("DELETE  FROM Magasin where [code Magasin] = '" + listBox10.Text + "' ", Con_HP_Four);
            MessageBox.Show("Vous avez supprimé le code magasin :  <" + LBL + " >");
            CMD_HP_FOUR.ExecuteNonQuery();
            Con_HP_Four.Close();
            Selection_Magasins_Exclus(); 
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                SqlConnection Con_HP_Four = new SqlConnection("Data Source=WAYNAV;Initial Catalog=Reporting-Hp-Config;Integrated Security=True");
                Con_HP_Four.Open();
                int Count_LBX1 = listBox2.Items.Count - 1;
                for (int i = Count_LBX1; i >= 0; i--)
                {
                    if (listBox2.GetSelected(i))
                    {
                        string LBL = listBox2.Text;
                        SqlCommand CMD_HP_FOUR = new SqlCommand("insert into Magasin ( [Code Magasin], [Magasin]) values ('" + listBox2.Text + "','NULL') ", Con_HP_Four);
                        CMD_HP_FOUR.ExecuteNonQuery();
                        MessageBox.Show("Vous avez exclu le code magasin : < " + LBL + " >");



                    }
                }
                Con_HP_Four.Close();
                Selection_Magasins_Exclus();
            }
            catch
            {
                string LBL = listBox2.Text;
                MessageBox.Show("Le code magasin : < " + LBL + " > existe dèja dans la liste des magasins exclus !");
            }
        }
    }
}

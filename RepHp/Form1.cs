using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevComponents.DotNetBar;


namespace RepHp
{
    public partial class Form1 : Office2007Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void nouveauReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Report Rep = new Report();
            Rep.MdiParent = this;
            Rep.Show();
        }

       

        private void historiqueToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Historique HIS = new Historique();
            HIS.MdiParent = this;
            HIS.Show();
        }

        private void lancerTestToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void paramètresToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Parametrage PRM = new Parametrage();
            PRM.MdiParent = this;
            PRM.Show();
        }

        private void quitterToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void statistiquesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Statistiques STQ = new Statistiques();
            STQ.MdiParent = this;
            STQ.Show();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void planifierReportToolStripMenuItem_Click(object sender, EventArgs e)
        {

            PLAN STQ = new PLAN();
            STQ.MdiParent = this;
            STQ.Show();
        }

        private void administrationToolStripMenuItem_Click(object sender, EventArgs e)
        {

            Administration ADM = new Administration();
            ADM.MdiParent = this;
            ADM.Show();
        }

        private void aProposToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Apropos AP = new Apropos();
            AP.MdiParent = this;
            AP.Show();
        }

        private void incidentsEtatsHPToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Incidents INC = new Incidents();
            INC.MdiParent = this;
            INC.Show();
        }

       


        
      
    }
}
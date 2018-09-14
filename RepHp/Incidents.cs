using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevComponents.DotNetBar;
using System.IO;

namespace RepHp
{
    public partial class Incidents : Office2007Form
    {
        public Incidents()
        {
            InitializeComponent();
        }

        private void Incidents_Load(object sender, EventArgs e)
        {
            // TODO: cette ligne de code charge les données dans la table '_Reporting_Hp_StorageFilesDataSet5.Incidents_HP'. Vous pouvez la déplacer ou la supprimer selon vos besoins.
            this.incidents_HPTableAdapter.Fill(this._Reporting_Hp_StorageFilesDataSet5.Incidents_HP);

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
        private void buttonX1_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls) |*.xls";
            sfd.FileName = "Incidents_Etats_HP.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                ToCsV(dataGridView1, sfd.FileName);
                MessageBox.Show("Le fichier " + sfd.FileName + " a été créé avec succès !");
            }
        }
    }
}

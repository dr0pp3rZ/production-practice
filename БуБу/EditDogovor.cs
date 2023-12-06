using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Учет_договоров_страхования
{
    public partial class EditDogovor : Form
    {
        public EditDogovor()
        {
            InitializeComponent();
        }

        private void EditDogovor_Load(object sender, EventArgs e)
        {
            FillClient();
            FillAgent();
        }

        private void FillAgent()
        {
            string SQL = "SELECT * FROM [Страховые агенты]";

            OleDbConnection conn = new OleDbConnection(Properties.Settings.Default.ConnStr);
            conn.Open();

            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(SQL, conn);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);

            dataGridView2.DataSource = ds.Tables[0].DefaultView;

            conn.Close();
        }

        private void FillClient()
        {
            string SQL = "SELECT * FROM Клиенты";

            OleDbConnection conn = new OleDbConnection(Properties.Settings.Default.ConnStr);
            conn.Open();

            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(SQL, conn);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);

            dataGridView1.DataSource = ds.Tables[0].DefaultView;

            conn.Close();
        }

        private void dataGridView1_Click(object sender, EventArgs e)
        {
            textBoxClientId.Text = dataGridView1[0, dataGridView1.CurrentRow.Index].Value.ToString();
        }

        private void dataGridView2_Click(object sender, EventArgs e)
        {
            textBoxAgentID.Text = dataGridView2[0, dataGridView2.CurrentRow.Index].Value.ToString();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBoxClientId.Text.Replace(" ", "").Length > 0 && textBoxSumma.Text.Replace(" ", "").Length > 0 && textBoxTarif.Text.Replace(" ", "").Length > 0)
            {
                double procent = 0;
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1[0, i].Value.ToString() == textBoxClientId.Text)
                    {
                        procent = Convert.ToDouble(dataGridView1[5, i].Value.ToString());
                        break;
                    }
                }

                if ((Convert.ToDouble(textBoxTarif.Text) - procent) <= 0)
                {
                    textBoxPrem.Text = (Convert.ToDouble(textBoxSumma.Text) * (1)).ToString();
                }
                else
                {
                    textBoxPrem.Text = (Convert.ToDouble(textBoxSumma.Text) * (Convert.ToDouble(textBoxTarif.Text) - procent)).ToString();
                }
            }
            else
            {
                textBoxPrem.Text = "0";
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBoxKomis.Text = (Convert.ToDouble(textBoxPrem.Text) * 0.1).ToString();
        }
    }
}

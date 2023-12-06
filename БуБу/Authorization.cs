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
    public partial class Authorization : Form
    {
        public Authorization()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (OleDbConnection connection = new OleDbConnection())
            {
                connection.ConnectionString = Properties.Settings.Default.ConnStr;
                connection.Open();

                string text = "SELECT * FROM [Страховые агенты] WHERE [Номер телефона] = @Номер AND Пароль = @Пароль";
                OleDbCommand command = new OleDbCommand(text, connection);
                command.Parameters.AddWithValue("@Номер", textBoxLogin.Text);
                command.Parameters.AddWithValue("@Пароль", textBoxPassword.Text);

                if (command.ExecuteScalar() == null)
                {
                    MessageBox.Show("Неправильно введен Логин или Пароль!");
                }
                else
                {
                    Form1 poliklinika = new Form1();
                    poliklinika.Show();

                    this.Hide();
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Registration registration= new Registration();
            registration.Show();
        }
    }
}

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
    public partial class Registration : Form
    {
        public Registration()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBoxFIO.Text.Replace(" ", "").Length > 0 && textBoxPassword.Text.Replace(" ", "").Length > 0 && textBoxNumber.Text.Replace(" ", "").Length > 0)
            {
                OleDbConnection connection = new OleDbConnection(Properties.Settings.Default.ConnStr);
                connection.Open();

                string text = "INSERT INTO [Страховые агенты] (ФИО, [Дата рождения], [Адрес], [Номер телефона], Пароль)" +
                    " VALUES (@ФИО, @Дата, @Адрес, @Номер, @Пароль)";
                OleDbCommand command = new OleDbCommand(text, connection);
                command.Parameters.AddWithValue("@ФИО", textBoxFIO.Text);
                command.Parameters.AddWithValue("@Дата", dateTimePicker1.Value.Date);
                command.Parameters.AddWithValue("@Адрес", textBoxAdres.Text);
                command.Parameters.AddWithValue("@Номер", textBoxNumber.Text);
                command.Parameters.AddWithValue("@Пароль", textBoxPassword.Text);
                command.ExecuteNonQuery();

                connection.Close();
                MessageBox.Show("Аккаунт создан");
                this.Close();
            }
            else
            {
                MessageBox.Show("Заполнены не все поля");
            }
        }
    }
}

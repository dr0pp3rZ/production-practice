using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Учет_договоров_страхования
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private string _path;
        
 
        
        BackgroundWorker bw = new BackgroundWorker
        {
            WorkerReportsProgress = true,
            WorkerSupportsCancellation = true
        };
        private void Form1_Load(object sender, EventArgs e)
        {
            FillClient();
            FillDogovor();
            FillAgent();
        }

        private void FillClient()
        {
            string SQL = "SELECT * FROM Клиенты";

            OleDbConnection conn = new OleDbConnection(Properties.Settings.Default.ConnStr);
            conn.Open();

            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(SQL, conn);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);

            dataGridViewClient.DataSource = ds.Tables[0].DefaultView;

            conn.Close();
        }

        private void FillDogovor()
        {
            string SQL = "SELECT * FROM Договоры";

            OleDbConnection conn = new OleDbConnection(Properties.Settings.Default.ConnStr);
            conn.Open();

            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(SQL, conn);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);

            dataGridViewDogovor.DataSource = ds.Tables[0].DefaultView;

            conn.Close();
        }

        private void FillAgent()
        {
            string SQL = "SELECT * FROM [Страховые агенты]";

            OleDbConnection conn = new OleDbConnection(Properties.Settings.Default.ConnStr);
            conn.Open();

            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(SQL, conn);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);

            dataGridViewAgent.DataSource = ds.Tables[0].DefaultView;

            conn.Close();
        }

        // Добавить клиента
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            AddClient form = new AddClient();
            form.ShowDialog();
            if (form.DialogResult == DialogResult.OK)
            {
                OleDbConnection connection = new OleDbConnection(Properties.Settings.Default.ConnStr);
                connection.Open();

                string text = "INSERT INTO Клиенты ([ФИО], [Дата рождения], [Адрес], [Номер телефона], [Индивидуальная скидка]) VALUES(@ФИО, @Дата, @Адрес, @Номер, @скидка)";
                OleDbCommand command = new OleDbCommand(text, connection);

                command.Parameters.AddWithValue("@ФИО", form.textBoxFIO.Text);
                command.Parameters.AddWithValue("@Дата", form.dateTimePicker1.Value.Date);
                command.Parameters.AddWithValue("@Адрес", form.textBoxAdres.Text);
                command.Parameters.AddWithValue("@Номер", form.textBoxNumber.Text);
                command.Parameters.AddWithValue("@скидка", form.textBoxSkidka.Text);

                command.ExecuteNonQuery();
                connection.Close();

                MessageBox.Show("Запись добавлена!");

                FillClient();
            }    
        }

        // Добавить договор
        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            AddDogovor form = new AddDogovor();
            form.ShowDialog();
            if (form.DialogResult == DialogResult.OK)
            {
                OleDbConnection connection = new OleDbConnection(Properties.Settings.Default.ConnStr);
                connection.Open();

                string text = "INSERT INTO [Договоры] ([Клиент], [Вид страхования], [Тариф], [Сумма страхования], [Страховая премия], [Дата заключения договора], [Страховой агент]) VALUES(@Клиент, @Вид, @Тариф, @Сумма, @премия, @Дата, @агент)";
                OleDbCommand command = new OleDbCommand(text, connection);

                command.Parameters.AddWithValue("@Клиент", form.textBoxClientId.Text);
                command.Parameters.AddWithValue("@Вид", form.textBoxVidStr.Text);
                command.Parameters.AddWithValue("@Тариф", form.textBoxTarif.Text);
                command.Parameters.AddWithValue("@Сумма", form.textBoxSumma.Text);
                command.Parameters.AddWithValue("@премия", form.textBoxPrem.Text);
                command.Parameters.AddWithValue("@Дата", form.dateTimePicker1.Value.Date);
                command.Parameters.AddWithValue("@агент", form.textBoxAgentID.Text);

                command.ExecuteNonQuery();
                connection.Close();

                MessageBox.Show("Запись добавлена!");

                FillDogovor();
            }
        }

        // Добавить агента
        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            AddAgent form = new AddAgent();
            form.ShowDialog();
            if (form.DialogResult == DialogResult.OK)
            {
                OleDbConnection connection = new OleDbConnection(Properties.Settings.Default.ConnStr);
                connection.Open();

                string text = "INSERT INTO [Страховые агенты] ([ФИО], [Дата рождения], [Адрес], [Номер телефона], Пароль) VALUES(@ФИО, @Дата, @Адрес, @Номер, @Пароль)";
                OleDbCommand command = new OleDbCommand(text, connection);

                command.Parameters.AddWithValue("@ФИО", form.textBoxFIO.Text);
                command.Parameters.AddWithValue("@Дата", form.dateTimePicker1.Value.Date);
                command.Parameters.AddWithValue("@Адрес", form.textBoxAdres.Text);
                command.Parameters.AddWithValue("@Номер", form.textBoxNumber.Text);
                command.Parameters.AddWithValue("@Пароль", form.textBoxPassword.Text);

                command.ExecuteNonQuery();
                connection.Close();

                MessageBox.Show("Запись добавлена!");

                FillAgent();
            }
        }


        // Редактировать клиента
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            EditClient form = new EditClient();

            int index = dataGridViewClient.CurrentRow.Index;
            string ID = dataGridViewClient[0, index].Value.ToString();
            form.labelID.Text += ID;
            form.textBoxFIO.Text = dataGridViewClient[1, index].Value.ToString();
            form.dateTimePicker1.Value = Convert.ToDateTime(dataGridViewClient[2, index].Value);
            form.textBoxAdres.Text = dataGridViewClient[3, index].Value.ToString();
            form.textBoxNumber.Text = dataGridViewClient[4, index].Value.ToString();           
            form.textBoxSkidka.Text = dataGridViewClient[5, index].Value.ToString();

            form.ShowDialog();
            if (form.DialogResult == DialogResult.OK)
            {
                OleDbConnection connection = new OleDbConnection(Properties.Settings.Default.ConnStr);
                connection.Open();  

                string text = "UPDATE Клиенты SET [ФИО] = @ФИО, [Дата рождения] = @Дата, [Адрес] = @Адрес," +
                    " [Номер телефона] = @Номер, [Индивидуальная скидка] = @скидка WHERE Код = @Код";
                OleDbCommand command = new OleDbCommand(text, connection);

                command.Parameters.AddWithValue("@ФИО", form.textBoxFIO.Text);
                command.Parameters.AddWithValue("@Дата", form.dateTimePicker1.Value.Date.ToString());
                command.Parameters.AddWithValue("@Адрес", form.textBoxAdres.Text);
                command.Parameters.AddWithValue("@Номер", form.textBoxNumber.Text);
                command.Parameters.AddWithValue("@скидка", form.textBoxSkidka.Text);
                command.Parameters.AddWithValue("@Код", ID);

                command.ExecuteNonQuery();
                connection.Close();

                MessageBox.Show("Запись обновлена!");

                FillClient();
            }
        }

        // Редактировать договор
        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            EditDogovor form = new EditDogovor();

            int index = dataGridViewDogovor.CurrentRow.Index;
            string ID = dataGridViewDogovor[0, index].Value.ToString();
            form.labelID.Text += ID;
            form.textBoxClientId.Text = dataGridViewDogovor[1, index].Value.ToString();
            form.textBoxVidStr.Text = dataGridViewDogovor[2, index].Value.ToString();
            form.textBoxTarif.Text = dataGridViewDogovor[3, index].Value.ToString();
            form.textBoxSumma.Text = dataGridViewDogovor[4, index].Value.ToString();
            form.textBoxPrem.Text = dataGridViewDogovor[5, index].Value.ToString();
            form.dateTimePicker1.Value = Convert.ToDateTime(dataGridViewDogovor[6, index].Value);
            form.textBoxAgentID.Text = dataGridViewDogovor[7, index].Value.ToString();

            form.ShowDialog();
            if (form.DialogResult == DialogResult.OK)
            {
                OleDbConnection connection = new OleDbConnection(Properties.Settings.Default.ConnStr);
                connection.Open();

                string text = "UPDATE Договоры SET [Клиент] = @Клиент, [Вид страхования] = @Вид, [Тариф] = @Тариф," +
                    " [Сумма страхования] = @Сумма, [Страховая премия] = @премия, [Дата заключения договора] = @Дата," +
                    " [Страховой агент] = @агент WHERE Код = @Код";
                OleDbCommand command = new OleDbCommand(text, connection);

                command.Parameters.AddWithValue("@Клиент", form.textBoxClientId.Text);
                command.Parameters.AddWithValue("@Вид", form.textBoxVidStr.Text);
                command.Parameters.AddWithValue("@Тариф", form.textBoxTarif.Text);
                command.Parameters.AddWithValue("@Сумма", form.textBoxSumma.Text);
                command.Parameters.AddWithValue("@премия", form.textBoxPrem.Text);
                command.Parameters.AddWithValue("@Дата", form.dateTimePicker1.Value.Date.ToString());
                command.Parameters.AddWithValue("@агент", form.textBoxAgentID.Text);
                command.Parameters.AddWithValue("@Код", ID);

                command.ExecuteNonQuery();
                connection.Close();

                MessageBox.Show("Запись обновлена!");

                FillDogovor();
            }
        }

        // Редактировать агента
        private void toolStripButton8_Click(object sender, EventArgs e)
        {
            EditAgent form = new EditAgent();

            int index = dataGridViewAgent.CurrentRow.Index;
            string ID = dataGridViewAgent[0, index].Value.ToString();
            form.labelID.Text += ID;
            form.textBoxFIO.Text = dataGridViewAgent[1, index].Value.ToString();
            form.dateTimePicker1.Value = Convert.ToDateTime(dataGridViewAgent[2, index].Value);
            form.textBoxAdres.Text = dataGridViewAgent[3, index].Value.ToString();
            form.textBoxNumber.Text = dataGridViewAgent[4, index].Value.ToString();
            form.textBoxPassword.Text = dataGridViewAgent[5, index].Value.ToString();

            form.ShowDialog();
            if (form.DialogResult == DialogResult.OK)
            {
                OleDbConnection connection = new OleDbConnection(Properties.Settings.Default.ConnStr);
                connection.Open();

                string text = "UPDATE [Страховые агенты] SET [ФИО] = @ФИО, [Дата рождения] = @Дата, [Адрес] = @Адрес," +
                    " [Номер телефона] = @Номер, Пароль = @Пароль WHERE Код = @Код";
                OleDbCommand command = new OleDbCommand(text, connection);

                command.Parameters.AddWithValue("@ФИО", form.textBoxFIO.Text);
                command.Parameters.AddWithValue("@Дата", form.dateTimePicker1.Value.Date.ToString());
                command.Parameters.AddWithValue("@Адрес", form.textBoxAdres.Text);
                command.Parameters.AddWithValue("@Номер", form.textBoxNumber.Text);
                command.Parameters.AddWithValue("@Пароль", form.textBoxPassword.Text);
                command.Parameters.AddWithValue("@Код", ID);

                command.ExecuteNonQuery();
                connection.Close();

                MessageBox.Show("Запись обновлена!");

                FillAgent();
            }
        }


        // Удалить клиента
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            OleDbConnection connection = new OleDbConnection(Properties.Settings.Default.ConnStr);
            connection.Open();

            string text = "DELETE FROM Клиенты WHERE [Код] = @Код";

            if (MessageBox.Show($"Удалить клиента {dataGridViewClient[0, dataGridViewClient.CurrentRow.Index].Value}, {dataGridViewClient[1, dataGridViewClient.CurrentRow.Index].Value}?", "Удаление клиента", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                OleDbCommand command = new OleDbCommand(text, connection);
                command.Parameters.AddWithValue("@Код", dataGridViewClient[0, dataGridViewClient.CurrentRow.Index].Value);
                command.ExecuteNonQuery();
                MessageBox.Show("Клиент удален");
            }

            connection.Close();
            FillClient();
            FillDogovor();
        }

        // Удалить договор
        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            OleDbConnection connection = new OleDbConnection(Properties.Settings.Default.ConnStr);
            connection.Open();

            string text = "DELETE FROM Договоры WHERE [Код] = @Код";

            if(MessageBox.Show($"Удалить договор {dataGridViewDogovor[0, dataGridViewDogovor.CurrentRow.Index].Value}, {dataGridViewDogovor[1, dataGridViewDogovor.CurrentRow.Index].Value}, {dataGridViewDogovor[3, dataGridViewDogovor.CurrentRow.Index].Value}?", "Удаление договора", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                OleDbCommand command = new OleDbCommand(text, connection);
                command.Parameters.AddWithValue("@Код", dataGridViewDogovor[0, dataGridViewDogovor.CurrentRow.Index].Value);
                command.ExecuteNonQuery();
                MessageBox.Show("Договор удален");
            }

            connection.Close();
            FillDogovor();
        }

        // Удалить агента
        private void toolStripButton9_Click(object sender, EventArgs e)
        {
            OleDbConnection connection = new OleDbConnection(Properties.Settings.Default.ConnStr);
            connection.Open();

            string text = "DELETE FROM [Страховые агенты] WHERE [Код] = @Код";

            if (MessageBox.Show($"Удалить агента {dataGridViewAgent[0, dataGridViewAgent.CurrentRow.Index].Value}, {dataGridViewAgent[1, dataGridViewAgent.CurrentRow.Index].Value}?", "Удаление договора", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                OleDbCommand command = new OleDbCommand(text, connection);
                command.Parameters.AddWithValue("@Код", dataGridViewAgent[0, dataGridViewAgent.CurrentRow.Index].Value);
                command.ExecuteNonQuery();
                MessageBox.Show("Агент удален");
            }

            connection.Close();
            FillAgent();
        }


        // Поиск клиента
        private void toolStripButton10_Click(object sender, EventArgs e)
        {
            OleDbConnection conn = new OleDbConnection(Properties.Settings.Default.ConnStr);

            string sqlCommand = "SELECT * FROM Клиенты WHERE ";
            switch (toolStripComboBox1.SelectedIndex)
            {
                case 0:
                    sqlCommand += $"Код LIKE \'%{toolStripTextBox1.Text}%\'";
                    break;
                case 1:
                    sqlCommand += $"ФИО LIKE \'%{toolStripTextBox1.Text}%\'";
                    break;
                case 2:
                    sqlCommand += $"[Дата рождения] LIKE \'%{toolStripTextBox1.Text}%\'";
                    break;
                case 3:
                    sqlCommand += $"Адрес LIKE \'%{toolStripTextBox1.Text}%\'";
                    break;
                case 4:
                    sqlCommand += $"[Номер телефона] LIKE \'%{toolStripTextBox1.Text}%\'";
                    break;
                case 5:
                    sqlCommand += $"[Индивидуальная скидка] LIKE \'%{toolStripTextBox1.Text}%\'";
                    break;
                default:
                    MessageBox.Show("Не выбран параметр для поиска!", "Внимание!");
                    return;
            }

            conn.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(sqlCommand, conn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridViewClient.DataSource = dt;
            conn.Close();
        }

        // Поиск договора
        private void toolStripButton11_Click(object sender, EventArgs e)
        {
            OleDbConnection conn = new OleDbConnection(Properties.Settings.Default.ConnStr);

            string sqlCommand = "SELECT * FROM [Договоры] WHERE ";
            switch (toolStripComboBox2.SelectedIndex)
            {
                case 0:
                    sqlCommand += $"Код LIKE \'%{toolStripTextBox2.Text}%\'";
                    break;
                case 1:
                    sqlCommand += $"Клиент LIKE \'%{toolStripTextBox2.Text}%\'";
                    break;
                case 2:
                    sqlCommand += $"[Вид страхования] LIKE \'%{toolStripTextBox2.Text}%\'";
                    break;
                case 3:
                    sqlCommand += $"[Тариф] LIKE \'%{toolStripTextBox2.Text}%\'";
                    break;
                case 4:
                    sqlCommand += $"[Сумма страхования] LIKE \'%{toolStripTextBox2.Text}%\'";
                    break;
                case 5:
                    sqlCommand += $"[Страховая премия] LIKE \'%{toolStripTextBox2.Text}%\'";
                    break;
                case 6:
                    sqlCommand += $"[Дата заключения договора] LIKE \'%{toolStripTextBox2.Text}%\'";
                    break;
                case 7:
                    sqlCommand += $"[Страховой агент] LIKE \'%{toolStripTextBox2.Text}%\'";
                    break;
                default:
                    MessageBox.Show("Не выбран параметр для поиска!", "Внимание!");
                    return;
            }

            conn.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(sqlCommand, conn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridViewDogovor.DataSource = dt;
            conn.Close();
        }

        // Поиск агента
        private void toolStripButton12_Click(object sender, EventArgs e)
        {
            OleDbConnection conn = new OleDbConnection(Properties.Settings.Default.ConnStr);

            string sqlCommand = "SELECT * FROM [Страховые агенты] WHERE ";
            switch (toolStripComboBox3.SelectedIndex)
            {
                case 0:
                    sqlCommand += $"Код LIKE \'%{toolStripTextBox3.Text}%\'";
                    break;
                case 1:
                    sqlCommand += $"ФИО LIKE \'%{toolStripTextBox3.Text}%\'";
                    break;
                case 2:
                    sqlCommand += $"[Дата рождения] LIKE \'%{toolStripTextBox3.Text}%\'";
                    break;
                case 3:
                    sqlCommand += $"Адрес LIKE \'%{toolStripTextBox3.Text}%\'";
                    break;
                case 4:
                    sqlCommand += $"[Номер телефона] LIKE \'%{toolStripTextBox3.Text}%\'";
                    break;
                default:
                    MessageBox.Show("Не выбран параметр для поиска!", "Внимание!");
                    return;
            }

            conn.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(sqlCommand, conn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridViewAgent.DataSource = dt;
            conn.Close();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void dataGridViewClient_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog od = new OpenFileDialog();
            od.Filter = "Excell|*.xls;*.xlsx;";
            od.FileName = "sd_tickets.xlsx";
            DialogResult dr = od.ShowDialog();
            if (dr == DialogResult.Abort)
                return;
            if (dr == DialogResult.Cancel)
                return;
            textBox1.Text = od.FileName.ToString();
            button2.Visible = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button2.Visible = false;
            _path = textBox1.Text;
            if (textBox1.Text == "" || !textBox1.Text.Contains("sd_tickets.xlsx"))
            {
                MessageBox.Show("Please Browse sd_tickets.xlsx to upload", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                textBox1.Text = "";
                button2.Visible = false;
                return;
            }
            if (bw.IsBusy)
            {
                return;
            }

            System.Diagnostics.Stopwatch sWatch = new System.Diagnostics.Stopwatch();
            bw.DoWork += (bwSender, bwArg) =>
            {
                //what happens here must not touch the form
                //as it's in a different thread
                sWatch.Start();
             
            };

            bw.ProgressChanged += (bwSender, bwArg) =>
            {
                //update progress bars here
            };

            bw.RunWorkerCompleted += (bwSender, bwArg) =>
            {
                //now you're back in the UI thread you can update the form
                //remember to dispose of bw now

                sWatch.Stop();

                //work is done, no need for the stop button now...

                textBox1.Text = "";
                button1.Enabled = true;
                
                bw.Dispose();
            };

            //lets allow the user to click stop

          
            button1.Enabled = false;

            //Starts the actual work - triggerrs the "DoWork" event
            bw.RunWorkerAsync();

            //InsertExcelRecords();
        }
    }
}

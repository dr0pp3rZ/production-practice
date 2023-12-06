using System.ComponentModel;
using System.Data;
using System.Drawing.Text;
using System.Xml;
using ExcelDataReader;
using MaterialSkin.Controls;
using ClosedXML.Excel;
using System.Windows.Forms;
using OfficeOpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace obzor
{
    public partial class Form1 : Form
    {
        private string fileName = string.Empty;
        private DataTableCollection tableCollection = null;
        private DataTable originalDataTable; // Переменная для хранения исходных данных
        private ExcelPackage excelPackage;
        private ExcelWorksheet worksheet;

        public Form1()
        {





            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult res = openFileDialog1.ShowDialog();

                if (res == DialogResult.OK)
                {
                    fileName = openFileDialog1.FileName;

                    Text = fileName;

                    OpenExcelFile(fileName);
                }
                else
                {
                    throw new Exception("Файл не выбран!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void OpenExcelFile(string path)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);


            IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);

            DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });

            tableCollection = db.Tables;

            toolStripComboBox1.Items.Clear();

            foreach (DataTable tabe in tableCollection)
            {
                toolStripComboBox1.Items.Add(tabe.TableName);
            }
            toolStripComboBox1.SelectedIndex = 0;
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable table = tableCollection[Convert.ToString(toolStripComboBox1.SelectedItem)];

            dataGridView1.DataSource = table;
        }

        // Метод загрузки данных из DataGridView и сохранения их в переменной originalDataTable
        private void LoadDataIntoDataTable()
        {
            originalDataTable = ((DataTable)dataGridView1.DataSource).Copy();
        }

        private void toolStripMenuEditor_Click(object sender, EventArgs e)
        {
            if (dataGridView1.ReadOnly == true)
            {
                dataGridView1.ReadOnly = false;
                toolStripMenuEditor.Text = "Выйти из режима редактирования";
            }
            else
            {
                DialogResult result = MessageBox.Show("Вы точно хотите сохранить внесённые изменения?", "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    // Код для сохранения изменений
                    // Например, сохранение данных в Excel файл или базу данных
                    // dataGridView1.EndEdit();
                    // (здесь код для сохранения данных)

                    // Перезапись originalDataTable после сохранения изменений
                    LoadDataIntoDataTable();
                }
                else
                {
                    // Отмена изменений
                    if (originalDataTable != null)
                    {
                        ((DataTable)dataGridView1.DataSource).Clear();
                        foreach (DataRow row in originalDataTable.Rows)
                        {
                            ((DataTable)dataGridView1.DataSource).ImportRow(row);
                        }
                    }
                }

                dataGridView1.ReadOnly = true;
                toolStripMenuEditor.Text = "Изменить";
            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (excelPackage != null && worksheet != null)
            {
                // Определите номер следующей строки для вставки
                int nextRow = worksheet.Dimension.End.Row + 1;

                // Получите доступ к данным из DataGridView
                DataGridViewRow newDataGridViewRow = new DataGridViewRow();
                newDataGridViewRow.CreateCells(dataGridView1);

                // Добавьте вашу логику здесь для заполнения новой строки данными из DataGridView
                // Например, вы можете пройти по ячейкам DataGridView и заполнить новую строку данными

                // После того, как вы заполните новую строку данными, вставьте её в Excel
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    worksheet.Cells[nextRow, i + 1].Value = newDataGridViewRow.Cells[i].Value;
                }

                // Сохраните изменения в файле Excel
                excelPackage.Save();

                MessageBox.Show("Новая строка добавлена в файл Excel.");
            }
            else
            {
                MessageBox.Show("Файл Excel не загружен.");
            }

        }
    }
}
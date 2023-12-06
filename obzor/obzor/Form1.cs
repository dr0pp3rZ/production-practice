using System.ComponentModel;
using System.Data;
using System.Drawing.Text;
using System.Xml;
using ExcelDataReader;
using MaterialSkin;
using MaterialSkin.Controls;
using OfficeOpenXml;
using System;
using System.IO;
using System.Windows.Forms;
using ClosedXML.Excel;



namespace obzor
{
    public partial class Form1 : MaterialForm
    {
        private string fileName = string.Empty;
        private DataTableCollection tableCollection = null;
        private readonly MaterialSkinManager materialSkinManager;
        public Form1()
        {
            InitializeComponent();

            // Инициализация materialSkinManager
            materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(
                Primary.Orange800, Primary.Orange900,
                Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE
            );



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

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            try
            {
                // Получаем выбранные ячейки
                DataGridViewSelectedCellCollection selectedCells = dataGridView1.SelectedCells;

                if (selectedCells.Count > 0)
                {
                    // Получаем индексы выбранных ячеек
                    int rowIndex = selectedCells[0].RowIndex;
                    int columnIndex = selectedCells[0].ColumnIndex;

                    // Получаем DataTable и удаляем ячейку
                    DataTable table = (DataTable)dataGridView1.DataSource;
                    table.Rows[rowIndex][columnIndex] = DBNull.Value;

                    // Обновляем DataGridView
                    dataGridView1.Refresh();
                }
                else
                {
                    MessageBox.Show("Выберите ячейку для удаления.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при удалении ячейки: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                // Завершаем редактирование ячейки и применяем все изменения
                dataGridView1.EndEdit();

                // Получаем DataTable из источника данных DataGridView
                DataTable table = (DataTable)dataGridView1.DataSource;

                if (!string.IsNullOrEmpty(fileName))
                {
                    // Создаем экземпляр класса для работы с Excel
                    using (XLWorkbook workbook = new XLWorkbook())
                    {
                        // Создаем новый лист в Excel
                        var worksheet = workbook.Worksheets.Add("Sheet1");

                        // Заполняем лист данными из DataTable
                        worksheet.Cell(1, 1).InsertTable(table);

                        // Сохраняем Excel-файл в тот же самый файл
                        workbook.SaveAs(fileName);

                        MessageBox.Show("Данные сохранены успешно.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Файл не был открыт. Выберите файл перед сохранением.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при сохранении данных: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                // Завершаем редактирование ячейки и применяем все изменения
                dataGridView1.EndEdit();

                // Получаем DataTable из источника данных DataGridView
                DataTable table = (DataTable)dataGridView1.DataSource;

                // Создаем новый DataTable с той же структурой
                DataTable newTable = table.Clone();

                // Копируем данные из DataGridView в новый DataTable
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    DataRow newRow = newTable.NewRow();
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        newRow[i] = row.Cells[i].Value;
                    }
                    newTable.Rows.Add(newRow);
                }

                // Заменяем старый DataTable новым DataTable
                table.Clear();
                foreach (DataRow row in newTable.Rows)
                {
                    table.ImportRow(row);
                }

                MessageBox.Show("Данные обновлены успешно.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при обновлении данных: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void materialSwitch1_CheckedChanged(object sender, EventArgs e)
        {
            if (materialSwitch1.Checked)
            {
                materialSkinManager.Theme = MaterialSkinManager.Themes.DARK;
                materialSkinManager.ColorScheme = new ColorScheme(
                    Primary.Teal800, Primary.Teal900,
                    Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE
                );
            }
            else
            {
                materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
                materialSkinManager.ColorScheme = new ColorScheme(
                    Primary.Orange800, Primary.Orange900,
                    Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE
                );
            }
        }

        private void materialSlider1_Click(object sender, EventArgs e)
        {
            // Преобразуем значение слайдера в коэффициент масштабирования (от 0.1 до 2.0, например)
            float scaleValue = 0.1f + (float)materialSlider1.Value / 50.0f;

            // Применяем масштабирование к форме
            this.Scale(new SizeF(scaleValue, scaleValue));

            // Обновляем масштабирование и расположение контролов в форме
            foreach (Control control in this.Controls)
            {
                control.Left = (int)(control.Left * scaleValue);
                control.Top = (int)(control.Top * scaleValue);
                control.Width = (int)(control.Width * scaleValue);
                control.Height = (int)(control.Height * scaleValue);

                // Дополнительные настройки для определенных типов контролов
                if (control is TextBox)
                {
                    TextBox textBox = (TextBox)control;
                    textBox.Font = new Font(textBox.Font.FontFamily, textBox.Font.Size * scaleValue);
                }
                // Добавьте дополнительные настройки для других типов контролов, если необходимо
            }
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }
    }
}

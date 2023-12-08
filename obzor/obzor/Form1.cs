using ClosedXML.Excel;
using ExcelDataReader;
using MaterialSkin;
using MaterialSkin.Controls;
using System.Data;
using ColorScheme = MaterialSkin.ColorScheme;

namespace obzor
{
    public partial class Form1 : MaterialForm
    {
        private string fileName = string.Empty; // Переменная для хранения пути к файлу
        private DataTableCollection tableCollection = null; // Коллекция таблиц из Excel
        private DataTable originalDataTable; // Исходная таблица для хранения оригинальных данных
        private readonly MaterialSkinManager materialSkinManager; // Менеджер MaterialSkin

        public Form1()
        {
            InitializeComponent();

            // Инициализация MaterialSkinManager
            materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(
                Primary.Orange800, Primary.Orange900,
                Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE
            );
        }

        // Метод для открытия файла Excel
        private void OpenExcelFile(string path)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);

            IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);

            DataSet dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });

            tableCollection = dataSet.Tables;

            toolStripComboBox1.Items.Clear();

            // Заполнение выпадающего списка названиями листов Excel
            foreach (DataTable table in tableCollection)
            {
                toolStripComboBox1.Items.Add(table.TableName);
            }

            toolStripComboBox1.SelectedIndex = 0;
        }

        // Обработчик выбора листа из выпадающего списка
        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable table = tableCollection[Convert.ToString(toolStripComboBox1.SelectedItem)];

            dataGridView1.DataSource = table;
        }

        // Метод для загрузки данных в DataTable из DataGridView
        private void LoadDataIntoDataTable()
        {
            originalDataTable = ((DataTable)dataGridView1.DataSource).Copy();
        }

        // Метод изменения состояния кнопок в режиме редактирования
        private void UpdateButtonAvailability()
        {
            сохранитьToolStripMenuItem1.Enabled = dataGridView1.ReadOnly;
            обновитьToolStripMenuItem1.Enabled = dataGridView1.ReadOnly;
            открытьToolStripMenuItem.Enabled = dataGridView1.ReadOnly;

            if (dataGridView1.ReadOnly)
            {
                toolStripMenuEditor.Text = "Редактировать";
            }
            else
            {
                toolStripMenuEditor.Text = "Выйти из режима редактирования";
                LoadDataIntoDataTable();
            }
        }

        // Обработчик нажатия на кнопку "Редактировать"
        private void toolStripMenuEditor_Click(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(fileName))
                {
                    dataGridView1.ReadOnly = !dataGridView1.ReadOnly;
                    UpdateButtonAvailability();
                }
                else
                {
                    MessageBox.Show("Файл не был открыт. Выберите файл перед редактированием.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (InvalidOperationException ex)
            {

            }
        }

        // Обработчик нажатия на кнопку "Добавить"
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable sourceTable = dataGridView1.DataSource as DataTable;
                if (HasErrors())
                {
                    dataGridView1.EndEdit();
                    return;
                }
                else
                {
                    if (sourceTable != null)
                    {
                        DataTable newTable = sourceTable.Clone();

                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            DataRow newRow = newTable.NewRow();
                            for (int i = 0; i < dataGridView1.Columns.Count; i++)
                            {
                                object cellValue = row.Cells[i].Value;
                                newRow[i] = cellValue != null ? cellValue : DBNull.Value;
                            }
                            newTable.Rows.Add(newRow);
                        }

                        sourceTable.Clear();
                        foreach (DataRow row in newTable.Rows)
                        {
                            sourceTable.ImportRow(row);
                        }

                        MessageBox.Show("Строки успешно добавлены в конец таблицы.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                        MessageBox.Show("Файл не был открыт. Выберите файл перед добавлением данных.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при добавлении строк: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Обработчик нажатия на кнопку "Удалить"
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(fileName))
            {
                try
                {
                    DataGridViewSelectedCellCollection selectedCells = dataGridView1.SelectedCells;

                    if (selectedCells.Count > 0)
                    {
                        int rowIndex = selectedCells[0].RowIndex;
                        int columnIndex = selectedCells[0].ColumnIndex;

                        DataTable table = (DataTable)dataGridView1.DataSource;
                        table.Rows[rowIndex][columnIndex] = DBNull.Value;

                        dataGridView1.Refresh();
                        LoadDataIntoDataTable();
                    }
                    else
                    {
                        MessageBox.Show("Выберите ячейку для очистки.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка при очистке ячейки: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
                MessageBox.Show("Файл не был открыт. Выберите файл перед удалением данных.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        // Обработчик события изменения MaterialSwitch для изменения темы формы
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

            UpdateDataGridViewColors();
        }

        // Метод обновления цвета ячеек в DataGridView
        private void UpdateDataGridViewColors()
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    // Обновляем цвета фона и текста ячейки
                    cell.Style.BackColor = dataGridView1.DefaultCellStyle.BackColor;
                    cell.Style.ForeColor = dataGridView1.DefaultCellStyle.ForeColor.Darken(1);
                }
            }
        }

        // Обработчик нажатия на кнопку "Сохранить"
        private void сохранитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.EndEdit();

                DataTable table = (DataTable)dataGridView1.DataSource;

                if (!string.IsNullOrEmpty(fileName))
                {
                    using (XLWorkbook workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Лист1");

                        worksheet.Cell(1, 1).InsertTable(table);

                        workbook.SaveAs(fileName);

                        MessageBox.Show("Данные успешно сохранены.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        // Обработчик нажатия на кнопку "Обновить"
        private void обновитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(fileName))
                {
                    MessageBox.Show("Файл не открыт. Откройте файл перед обновлением данных.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                dataGridView1.EndEdit();

                DataTable table = (DataTable)dataGridView1.DataSource;

                if (table != null)
                {
                    DataTable newTable = table.Clone();


                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        DataRow newRow = newTable.NewRow();
                        for (int i = 0; i < dataGridView1.Columns.Count; i++)
                        {
                            if (row.Cells[i].Value != null)
                            {
                                newRow[i] = row.Cells[i].Value;
                            }
                        }
                        newTable.Rows.Add(newRow);
                    }

                    // Очищаем исходный DataTable и импортируем данные из нового DataTable
                    table.Clear();
                    foreach (DataRow row in newTable.Rows)
                    {
                        table.ImportRow(row);
                    }

                    MessageBox.Show("Данные обновлены успешно.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Исходная таблица равна null. Проверьте, что DataGridView связан с DataTable.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при обновлении данных: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void SaveChangesToFile()
        {
            if (originalDataTable != null && !originalDataTable.Equals(((DataTable)dataGridView1.DataSource)))
            {
                DialogResult result = MessageBox.Show("Есть несохраненные изменения. Хотите сохранить их?", "Предупреждение", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    try
                    {
                        using (XLWorkbook workbook = new XLWorkbook())
                        {
                            var worksheet = workbook.Worksheets.Add("Лист1");

                            worksheet.Cell(1, 1).InsertTable((DataTable)dataGridView1.DataSource);

                            workbook.SaveAs(fileName);

                            MessageBox.Show("Изменения сохранены успешно.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка при сохранении изменений: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        // Обработчик нажатия на кнопку "Загрузить"
        private void открытьToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            try
            {
                DialogResult res = openFileDialog1.ShowDialog();

                if (res == DialogResult.OK)
                {
                    string newFileName = openFileDialog1.FileName;

                    if (!string.Equals(newFileName, fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        if (!string.IsNullOrEmpty(fileName) && dataGridView1.DataSource != null)
                        {
                            SaveChangesToFile();
                        }

                        fileName = newFileName;
                        Text = fileName;

                        OpenExcelFile(fileName);
                    }
                    else
                    {
                        throw new Exception("Файл не выбран!");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Метод проверки наличия ошибок в ячейках
        private bool HasErrors()
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.ErrorText != null && cell.ErrorText.Trim() != "")
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        // Метод обработки ошибок, вызванных неправильно внесёнными данными в ячейки таблицы
        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            DataGridViewCell errorCell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
            errorCell.ErrorText = "Ошибка в ячейке";

            int displayedRow = e.RowIndex + 1;
            int displayedColumn = e.ColumnIndex + 1;

            MessageBox.Show($"Ошибка в ячейке. Номер строки: {displayedRow}, Номер столбца: {displayedColumn}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            e.ThrowException = false;
        }
    }
}
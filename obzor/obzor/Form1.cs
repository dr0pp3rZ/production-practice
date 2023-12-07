using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
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
            сохранитьToolStripMenuItem1.Enabled = !dataGridView1.ReadOnly;
            обновитьToolStripMenuItem1.Enabled = !dataGridView1.ReadOnly;
            открытьToolStripMenuItem.Enabled = !dataGridView1.ReadOnly;
        }

        // Обработчик нажатия на кнопку "Редактировать"
        private void toolStripMenuEditor_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(fileName))
            {
                UpdateButtonAvailability();
                if (dataGridView1.ReadOnly)
                {
                    dataGridView1.ReadOnly = false;
                    toolStripMenuEditor.Text = "Сохранить изм. в реж. редактирования";
                }
                else
                {
                    try
                    {
                        if (!HasDataGridViewErrors(out int errorRow, out int errorColumn))
                        {
                            // Проверка наличия выделенной ячейки перед завершением редактирования
                            if (dataGridView1.CurrentCell != null && dataGridView1.CurrentCell.IsInEditMode)
                            {
                                dataGridView1.EndEdit();
                            }

                            dataGridView1.ReadOnly = true;
                            toolStripMenuEditor.Text = "Редактировать";
                            LoadDataIntoDataTable();
                        }
                        else
                        {
                            MessageBox.Show($"Имеются ошибки в данных. Не удается сохранить изменения. " +
                                $"Номер строки: {errorRow}, Номер столбца: {errorColumn}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    catch (InvalidOperationException ex)
                    {
                        int errorRow = -1;
                        int errorColumn = -1;

                        if (dataGridView1.CurrentCell != null)
                        {
                            errorRow = dataGridView1.CurrentCell.RowIndex;
                            errorColumn = dataGridView1.CurrentCell.ColumnIndex;
                        }

                        MessageBox.Show($"Имеются ошибки в данных. Не удается сохранить изменения. " +
                            $"Номер строки: {errorRow}, " +
                            $"Номер столбца: {errorColumn}\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        // Метод проверки и вывода сообщения об ошибках
        private bool HasDataGridViewErrors(out int errorRow, out int errorColumn)
        {
            bool hasErrors = false;
            errorRow = -1;
            errorColumn = -1;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.ErrorText != string.Empty)
                    {
                        hasErrors = true;
                        errorRow = cell.RowIndex;
                        errorColumn = cell.ColumnIndex;
                        break;
                    }
                }

                if (hasErrors)
                {
                    break;
                }
            }

            return hasErrors;
        }

        // Обработчик нажатия на кнопку "Добавить"
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

            try
            {
                dataGridView1.EndEdit();

                DataTable sourceTable = dataGridView1.DataSource as DataTable;

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

                    MessageBox.Show("Строки добавлены успешно.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Исходная таблица не инициализирована.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        // Обработчик нажатия на кнопку "Загрузить"
        private void открытьToolStripMenuItem_Click_1(object sender, EventArgs e)
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

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            // Получить информацию об ошибке
            string columnName = dataGridView1.Columns[e.ColumnIndex].Name;
            string cellText = string.Empty;
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && e.RowIndex < dataGridView1.Rows.Count)
            {
                cellText = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? string.Empty;
            }

            // Выполнить необходимые действия для устранения ошибки
            if (e.Exception is FormatException)
            {
                // Формат данных не соответствует ожидаемому формату
                MessageBox.Show($"Значение в столбце \"{columnName}\" должно быть типа {dataGridView1.Columns[e.ColumnIndex].ValueType?.Name}");
            }
            else if (e.Exception is ConstraintException)
            {
                // Значение данных не соответствует ограничениям
                MessageBox.Show($"Значение в столбце \"{columnName}\" должно удовлетворять условию {e.Exception.Message}");
            }
            else
            {
                // Ошибка, которую необходимо обработать по умолчанию
                MessageBox.Show($"Ошибка в столбце \"{columnName}\". Значение: {cellText}. Тип ошибки: {e.Exception?.GetType()?.Name}");
            }
        }
    }
}
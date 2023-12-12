using ClosedXML.Excel;
using ExcelDataReader;
using MaterialSkin;
using System.Data;

namespace obzor.Methods
{
    internal class ButtonMethods
    {
        private readonly FileOperationMethods _fileOperations = new();
        Variables variables = new();
        // Метод для открытия файла Excel
        public void OpenExcelFile(DataGridView dataGridView1, string path, ToolStripComboBox toolStripComboBox1, ToolStripMenuItem ToolStripMenuEditor)
        {
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);
                IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
                DataSet dataSet = reader.AsDataSet(new ExcelDataSetConfiguration { ConfigureDataTable = (_) => new ExcelDataTableConfiguration { UseHeaderRow = true } });
                variables.tableCollection = dataSet.Tables;

                // Получение имен таблиц в виде списка
                List<string> tableNames = new();
                foreach (DataTable table in variables.tableCollection)
                {
                    tableNames.Add(table.TableName);
                }
                if (variables.tableCollection.Count > 0)
                {
                    dataGridView1.DataSource = variables.tableCollection[0];
                }
                // Обновление ToolStripComboBox с передачей списка имен таблиц
                FileOperationMethods.UpdateToolStripComboBox(tableNames, toolStripComboBox1, ToolStripMenuEditor);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка открытия файла: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Обработчик нажатия на кнопку "Загрузить"
        public void LoadFile(DataGridView dataGridView1, OpenFileDialog openFileDialog1, ToolStripComboBox toolStripComboBox1, ToolStripMenuItem ToolStripMenuEditor)
        {
            try
            {
                DialogResult res = openFileDialog1.ShowDialog();

                if (res == DialogResult.OK)
                {
                    string newFileName = openFileDialog1.FileName;

                    if (string.Equals(newFileName, variables.fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new Exception("Файл уже открыт!");
                    }

                    if (!string.IsNullOrEmpty(variables.fileName) && dataGridView1.DataSource != null)
                    {
                        _fileOperations.SaveChangesToFile(dataGridView1);
                    }

                    variables.fileName = newFileName;
                    Form.ActiveForm.Text = variables.fileName;

                    // Определение типа выбранного файла и выполнение соответствующих действий
                    string fileExtension = Path.GetExtension(variables.fileName);
                    switch (fileExtension.ToLower())
                    {
                        case ".xls":
                        case ".xlsx":
                            OpenExcelFile(dataGridView1, variables.fileName, toolStripComboBox1, ToolStripMenuEditor);
                            break;
                        case ".accdb":
                        case ".mdb":
                            //FileOperationMethods.OpenAccessFile(dataGridView1, variables.fileName, toolStripComboBox1, ToolStripMenuEditor);
                            break;

                        case ".mdf":
                            //FileOperationMethods.OpenSQLServerDatabase(dataGridView1, variables.fileName, toolStripComboBox1, ToolStripMenuEditor);
                            break;

                        default:
                            throw new Exception("Неподдерживаемый тип файла!");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ToolStripComboBoxSelectedIndexChanged(DataGridView dataGridView1, ToolStripComboBox toolStripComboBox1)
        {
            if (toolStripComboBox1.SelectedItem != null && variables.tableCollection != null)
            {
                string selectedTable = toolStripComboBox1.SelectedItem.ToString();
                if (variables.tableCollection.Contains(selectedTable))
                {
                    DataTable table = variables.tableCollection[selectedTable];
                    dataGridView1.DataSource = table;
                }
                else
                {
                    MessageBox.Show($"Таблица {selectedTable} не найдена.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // TODO: Дополнить код, для других баз данных
        // Обработчик нажатия на кнопку "Сохранить"
        public void SaveButton(DataGridView dataGridView1)
        {
            try
            {
                dataGridView1.EndEdit();

                DataTable table = (DataTable)dataGridView1.DataSource;

                if (!string.IsNullOrEmpty(variables.fileName))
                {
                    using XLWorkbook workbook = new();
                    var worksheet = workbook.Worksheets.Add("Лист1");

                    worksheet.Cell(1, 1).InsertTable(table);

                    workbook.SaveAs(variables.fileName);

                    MessageBox.Show("Данные успешно сохранены.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        // TODO: Дополнить код, для других баз данных
        // Обработчик нажатия на кнопку "Сохранить как"
        public void SaveAsNewFile(DataGridView dataGridView1, DataTable table, ToolStripComboBox toolStripComboBox1, ToolStripMenuItem ToolStripMenuEditor)
        {
            if (table is null)
            {
                throw new ArgumentNullException(nameof(table));
            }

            try
            {
                SaveFileDialog saveFileDialog = new()
                {
                    Filter = "Excel Files|*.xlsx;*.xls",
                    Title = "Сохранить как"
                };
                saveFileDialog.ShowDialog();

                if (!string.IsNullOrWhiteSpace(saveFileDialog.FileName))
                {
                    variables.fileName = saveFileDialog.FileName;
                    Form.ActiveForm.Text = variables.fileName;

                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Лист1");
                        DataTable dataTable = new("Лист1");

                        // Создание структуры таблицы (заголовков столбцов) из DataGridView
                        foreach (DataGridViewColumn column in dataGridView1.Columns)
                        {
                            dataTable.Columns.Add(column.HeaderText);
                        }

                        // Добавление данных из DataGridView в таблицу данных
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            DataRow dataRow = dataTable.NewRow();
                            for (int i = 0; i < dataGridView1.Columns.Count; i++)
                            {
                                dataRow[i] = row.Cells[i].Value ?? DBNull.Value;
                            }
                            dataTable.Rows.Add(dataRow);
                        }

                        worksheet.Cell(1, 1).InsertTable(dataTable);
                        workbook.SaveAs(variables.fileName);
                    }
                    // TODO: дополнить код проверкой какой тип файла сохраняется
                    // Открываем новый файл и выводим его содержимое в DataGridView
                    OpenExcelFile(dataGridView1, variables.fileName, toolStripComboBox1, ToolStripMenuEditor);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при создании нового файла: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Обработчик нажатия на кнопку "Добавить"
        public static void AddedButton(DataGridView dataGridView1)
        {
            try
            {
                if (FileOperationMethods.HasErrors(dataGridView1))
                {
                    dataGridView1.EndEdit();
                    return;
                }
                else
                {
                    if (dataGridView1.DataSource is DataTable sourceTable)
                    {
                        DataTable newTable = sourceTable.Clone();

                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            DataRow newRow = newTable.NewRow();
                            for (int i = 0; i < dataGridView1.Columns.Count; i++)
                            {
                                object cellValue = row.Cells[i].Value;
                                newRow[i] = cellValue ?? DBNull.Value;
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

        // Обработчик нажатия на кнопку "Редактировать"
        public void EditButton(DataGridView dataGridView1, ToolStripComboBox toolStripComboBox1, ToolStripMenuItem файлToolStripMenuItem, ToolStripMenuItem редактироватьToolStripMenuItem)
        {
            try
            {
                if (!string.IsNullOrEmpty(variables.fileName))
                {
                    dataGridView1.ReadOnly = !dataGridView1.ReadOnly;

                    _fileOperations.UpdateButtonAvailability(dataGridView1, toolStripComboBox1, файлToolStripMenuItem, редактироватьToolStripMenuItem);
                }
                else
                {
                    MessageBox.Show("Файл не был открыт. Выберите файл перед редактированием.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (InvalidOperationException)
            {

            }
        }

        // Обработчик нажатия на кнопку "Обновить"
        public void UpdateButton(DataGridView dataGridView1)
        {
            try
            {
                if (string.IsNullOrEmpty(variables.fileName))
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


        // Обработчик нажатия на кнопку "Удалить"
        public void ClearSelectedCell(DataGridView dataGridView1, string fileName)
        {
            if (!string.IsNullOrEmpty(variables.fileName))
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
                        _fileOperations.LoadDataIntoDataTable(dataGridView1);
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
    }
}
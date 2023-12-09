using ClosedXML.Excel;
using ExcelDataReader;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace obzor.Methods
{
    internal class FileOperationMethods
    {
        private DataTableCollection? tableCollection = null; // Коллекция таблиц из баз данных
        private DataTable? originalDataTable; // Исходная таблица для хранения оригинальных данных
        private readonly string fileName = string.Empty; // Переменная для хранения пути к файлу

        // Метод для открытия файла Excel
        public void OpenExcelFile(DataGridView dataGridView1, string path, ToolStripComboBox toolStripComboBox1, ToolStripMenuItem ToolStripMenuEditor)
        {
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);
                IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
                DataSet dataSet = reader.AsDataSet(new ExcelDataSetConfiguration { ConfigureDataTable = (_) => new ExcelDataTableConfiguration { UseHeaderRow = true } });
                tableCollection = dataSet.Tables;

                // Получение имен таблиц в виде списка
                List<string> tableNames = new();
                foreach (DataTable table in tableCollection)
                {
                    tableNames.Add(table.TableName);
                }
                if (tableCollection.Count > 0)
                {
                    dataGridView1.DataSource = tableCollection[0];
                }
                // Обновление ToolStripComboBox с передачей списка имен таблиц
                UpdateToolStripComboBox(tableNames, toolStripComboBox1, ToolStripMenuEditor);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка открытия файла: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // TODO: Доработать вывод и редактирование данных в бд access
        // Метод для открытия файла Access
        public static void OpenAccessFile(DataGridView dataGridView1, string path, ToolStripComboBox toolStripComboBox1, ToolStripMenuItem ToolStripMenuEditor)
        {
            try
            {
                string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={path};";
                using OleDbConnection connection = new(connectionString);
                connection.Open();

                DataTable schemaTable = connection.GetSchema("Tables");
                List<string> tableNames = new();

                foreach (DataRow row in schemaTable.Rows)
                {
                    string tableName = row["TABLE_NAME"].ToString();
                    tableNames.Add(tableName);
                }

                if (tableNames.Count > 0)
                {
                    string firstTableName = tableNames[0];
                    string query = $"SELECT * FROM [{firstTableName}]";
                    OleDbCommand command = new(query, connection);

                    OleDbDataAdapter adapter = new(command);
                    DataTable dataTable = new();
                    adapter.Fill(dataTable);

                    dataGridView1.DataSource = dataTable;
                    UpdateToolStripComboBox(tableNames, toolStripComboBox1, ToolStripMenuEditor);
                }
                else
                {
                    MessageBox.Show("База данных не содержит таблиц.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка открытия файла Access: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // TODO: Доработать вывод и редактирование данных в бд sql server
        // Метод для открытия файла SQLServer
        public static void OpenSQLServerDatabase(DataGridView dataGridView1, string connectionString, ToolStripComboBox toolStripComboBox1, ToolStripMenuItem ToolStripMenuEditor)
        {
            try
            {
                using SqlConnection connection = new(connectionString);
                connection.Open();
                DataTable schemaTable = connection.GetSchema("Tables");

                List<string> tableNames = new();
                foreach (DataRow row in schemaTable.Rows)
                {
                    string tableName = row["TABLE_NAME"].ToString();
                    tableNames.Add(tableName);
                }

                if (tableNames.Count > 0)
                {
                    string firstTableName = tableNames[0];
                    string query = $"SELECT * FROM [{firstTableName}]";
                    SqlCommand command = new(query, connection);

                    SqlDataAdapter adapter = new(command);
                    DataTable dataTable = new();
                    adapter.Fill(dataTable);

                    dataGridView1.DataSource = dataTable;
                    UpdateToolStripComboBox(tableNames, toolStripComboBox1, ToolStripMenuEditor);
                }
                else
                {
                    MessageBox.Show("База данных не содержит таблиц.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка подключения к базе данных MS SQL Server: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // TODO: Починить обновление данных при выборе другого листа
        // Проблема заключается в том, что ToolStripComboBoxSelectedIndexChanged и LoadFile находятся в разных файлах и возможно из-за tableCollection и originalDataTable
        // Обработчик выбора листа из выпадающего списка
        public void ToolStripComboBoxSelectedIndexChanged(DataGridView dataGridView1, ToolStripComboBox toolStripComboBox1)
        {
            if (toolStripComboBox1.SelectedItem != null && tableCollection != null)
            {
                string selectedTable = toolStripComboBox1.SelectedItem.ToString();
                if (tableCollection.Contains(selectedTable))
                {
                    DataTable table = tableCollection[selectedTable];
                    dataGridView1.DataSource = table;
                }
                else
                {
                    MessageBox.Show($"Таблица {selectedTable} не найдена.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // Метод обновления листов
        public static void UpdateToolStripComboBox(List<string> tableNames, ToolStripComboBox toolStripComboBox1, ToolStripMenuItem toolStripMenuEditor)
        {
            toolStripComboBox1.Items.Clear();
            if (tableNames != null && tableNames.Any())
            {
                foreach (string tableName in tableNames)
                {
                    toolStripComboBox1.Items.Add(tableName);
                }
                toolStripComboBox1.SelectedIndex = 0;
                toolStripComboBox1.Enabled = true;
                toolStripMenuEditor.Enabled = true;
            }
        }

        // Метод для загрузки данных в DataTable из DataGridView
        public void LoadDataIntoDataTable(DataGridView dataGridView1)
        {
            originalDataTable = ((DataTable)dataGridView1.DataSource).Copy();
        }

        // Метод изменения состояния кнопок в режиме редактирования
        public void UpdateButtonAvailability(DataGridView dataGridView1, ToolStripComboBox toolStripComboBox1, ToolStripMenuItem файлToolStripMenuItem, ToolStripMenuItem редактироватьToolStripMenuItem)
        {
            файлToolStripMenuItem.Enabled = dataGridView1.ReadOnly;
            toolStripComboBox1.Enabled = dataGridView1.ReadOnly;

            if (dataGridView1.ReadOnly)
            {
                редактироватьToolStripMenuItem.Text = "Редактировать";
            }
            else
            {
                редактироватьToolStripMenuItem.Text = "Выйти из режима редактирования";
                LoadDataIntoDataTable(dataGridView1);
            }
        }

        // Метод проверки наличия ошибок в ячейках
        public static bool HasErrors(DataGridView dataGridView1)
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
        public static void DataError(DataGridView dataGridView1, DataGridViewDataErrorEventArgs e)
        {
            DataGridViewCell errorCell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
            errorCell.ErrorText = "Ошибка в ячейке";

            int displayedRow = e.RowIndex + 1;
            int displayedColumn = e.ColumnIndex + 1;

            MessageBox.Show($"Ошибка в ячейке. Номер строки: {displayedRow}, Номер столбца: {displayedColumn}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            e.ThrowException = false;
        }

        // Метод для запроса сохранений перед выбором другого файла
        public void SaveChangesToFile(DataGridView dataGridView1)
        {
            if (originalDataTable != null && !originalDataTable.Equals((DataTable)dataGridView1.DataSource))
            {
                DialogResult result = MessageBox.Show("Есть несохраненные изменения. Хотите сохранить их?", "Предупреждение", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    try
                    {
                        using XLWorkbook workbook = new();
                        var worksheet = workbook.Worksheets.Add("Лист1");

                        worksheet.Cell(1, 1).InsertTable((DataTable)dataGridView1.DataSource);

                        workbook.SaveAs(fileName);

                        MessageBox.Show("Изменения сохранены успешно.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка при сохранении изменений: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
    }
}
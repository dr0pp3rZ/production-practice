using System;
using System.Data;
using System.Drawing.Text;
using System.IO;
using System.Windows.Forms;
using ExcelDataReader;
using ClosedXML.Excel;
using OfficeOpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using MaterialSkin.Controls;
using MaterialSkin;

namespace obzor
{
    public partial class Form1 : MaterialForm
    {
        private string fileName = string.Empty;
        private DataTableCollection tableCollection = null;
        private DataTable originalDataTable; // Исходная таблица для хранения оригинальных данных
        private ExcelPackage excelPackage;
        private ExcelWorksheet worksheet;
        private readonly MaterialSkinManager materialSkinManager;

        public Form1()
        {
            InitializeComponent();

            // Èíèöèàëèçàöèÿ materialSkinManager
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

        private void îòêðûòüToolStripMenuItem_Click(object sender, EventArgs e)
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
                    throw new Exception("Ôàéë íå âûáðàí!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Îøèáêà!", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

            foreach (DataTable table in tableCollection)
            {
                toolStripComboBox1.Items.Add(table.TableName);
            }

            toolStripComboBox1.SelectedIndex = 0;
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable table = tableCollection[Convert.ToString(toolStripComboBox1.SelectedItem)];

            dataGridView1.DataSource = table;
        }

        private void LoadDataIntoDataTable()
        {
            originalDataTable = ((DataTable)dataGridView1.DataSource).Copy();
        }

        private void toolStripMenuEditor_Click(object sender, EventArgs e)
        {
            if (dataGridView1.ReadOnly == true)
            {
                dataGridView1.ReadOnly = false;
                toolStripMenuEditor.Text = "Сохранить изм. в реж. редактирования";
            }
            else
            {
                DialogResult result = MessageBox.Show("Вы уверены, что хотите сохранить внесенные изменения?", "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    LoadDataIntoDataTable();
                }
                else
                {
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
                toolStripMenuEditor.Text = "Редактировать";
            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.EndEdit();

                DataTable table = (DataTable)dataGridView1.DataSource;

                DataTable newTable = table.Clone();

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    DataRow newRow = newTable.NewRow();
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        newRow[i] = row.Cells[i].Value;
                    }
                    newTable.Rows.Add(newRow);
                }

                table.Clear();
                foreach (DataRow row in newTable.Rows)
                {
                    table.ImportRow(row);
                }

                MessageBox.Show("Строки добавлены успешно.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при добавлении строк: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //if (excelPackage != null && worksheet != null)
            //{
            //    int nextRow = worksheet.Dimension.End.Row + 1;

            //    DataGridViewRow newDataGridViewRow = new DataGridViewRow();
            //    newDataGridViewRow.CreateCells(dataGridView1);

            //    for (int i = 0; i < dataGridView1.Columns.Count; i++)
            //    {
            //        worksheet.Cells[nextRow, i + 1].Value = newDataGridViewRow.Cells[i].Value;
            //    }

            //    excelPackage.Save();

            //    MessageBox.Show("Новая строка добавлена в файл Excel.");
            //}
            //else
            //{
            //    MessageBox.Show("Файл Excel не загружен.");
            //}
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
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

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
        }

        private void сохранитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                // Çàâåðøàåì ðåäàêòèðîâàíèå ÿ÷åéêè è ïðèìåíÿåì âñå èçìåíåíèÿ
                dataGridView1.EndEdit();

                // Ïîëó÷àåì DataTable èç èñòî÷íèêà äàííûõ DataGridView
                DataTable table = (DataTable)dataGridView1.DataSource;

                if (!string.IsNullOrEmpty(fileName))
                {
                    // Ñîçäàåì ýêçåìïëÿð êëàññà äëÿ ðàáîòû ñ Excel
                    using (XLWorkbook workbook = new XLWorkbook())
                    {
                        // Ñîçäàåì íîâûé ëèñò â Excel
                        var worksheet = workbook.Worksheets.Add("Sheet1");

                        // Çàïîëíÿåì ëèñò äàííûìè èç DataTable
                        worksheet.Cell(1, 1).InsertTable(table);

                        // Ñîõðàíÿåì Excel-ôàéë â òîò æå ñàìûé ôàéë
                        workbook.SaveAs(fileName);

                        MessageBox.Show("Äàííûå ñîõðàíåíû óñïåøíî.", "Óñïåõ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Ôàéë íå áûë îòêðûò. Âûáåðèòå ôàéë ïåðåä ñîõðàíåíèåì.", "Ïðåäóïðåæäåíèå", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Îøèáêà ïðè ñîõðàíåíèè äàííûõ: " + ex.Message, "Îøèáêà", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void обновитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                // Çàâåðøàåì ðåäàêòèðîâàíèå ÿ÷åéêè è ïðèìåíÿåì âñå èçìåíåíèÿ
                dataGridView1.EndEdit();

                // Ïîëó÷àåì DataTable èç èñòî÷íèêà äàííûõ DataGridView
                DataTable table = (DataTable)dataGridView1.DataSource;

                // Ñîçäàåì íîâûé DataTable ñ òîé æå ñòðóêòóðîé
                DataTable newTable = table.Clone();

                // Êîïèðóåì äàííûå èç DataGridView â íîâûé DataTable
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    DataRow newRow = newTable.NewRow();
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        newRow[i] = row.Cells[i].Value;
                    }
                    newTable.Rows.Add(newRow);
                }

                // Çàìåíÿåì ñòàðûé DataTable íîâûì DataTable
                table.Clear();
                foreach (DataRow row in newTable.Rows)
                {
                    table.ImportRow(row);
                }

                MessageBox.Show("Äàííûå îáíîâëåíû óñïåøíî.", "Óñïåõ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Îøèáêà ïðè îáíîâëåíèè äàííûõ: " + ex.Message, "Îøèáêà", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
    }
}

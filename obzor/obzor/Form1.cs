using System.ComponentModel;
using System.Data;
using System.Drawing.Text;
using System.Xml;
using ExcelDataReader;
using MaterialSkin;
using MaterialSkin.Controls;
using ClosedXML.Excel;
using System.Windows.Forms;
using OfficeOpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace obzor
{
    public partial class Form1 : MaterialForm
    {
        private string fileName = string.Empty;
        private DataTableCollection tableCollection = null;
        private DataTable originalDataTable; // Ïåðåìåííàÿ äëÿ õðàíåíèÿ èñõîäíûõ äàííûõ
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

        // Ìåòîä çàãðóçêè äàííûõ èç DataGridView è ñîõðàíåíèÿ èõ â ïåðåìåííîé originalDataTable
        private void LoadDataIntoDataTable()
        {
            originalDataTable = ((DataTable)dataGridView1.DataSource).Copy();
        }

        private void toolStripMenuEditor_Click(object sender, EventArgs e)
        {
            if (dataGridView1.ReadOnly == true)
            {
                dataGridView1.ReadOnly = false;
                toolStripMenuEditor.Text = "Âûéòè èç ðåæèìà ðåäàêòèðîâàíèÿ";
            }
            else
            {
                DialogResult result = MessageBox.Show("Âû òî÷íî õîòèòå ñîõðàíèòü âíåñ¸ííûå èçìåíåíèÿ?", "Ïîäòâåðæäåíèå", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    // Êîä äëÿ ñîõðàíåíèÿ èçìåíåíèé
                    // Íàïðèìåð, ñîõðàíåíèå äàííûõ â Excel ôàéë èëè áàçó äàííûõ
                    // dataGridView1.EndEdit();
                    // (çäåñü êîä äëÿ ñîõðàíåíèÿ äàííûõ)

                    // Ïåðåçàïèñü originalDataTable ïîñëå ñîõðàíåíèÿ èçìåíåíèé
                    LoadDataIntoDataTable();
                }
                else
                {
                    // Îòìåíà èçìåíåíèé
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
                toolStripMenuEditor.Text = "Èçìåíèòü";
            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (excelPackage != null && worksheet != null)
            {
                // Îïðåäåëèòå íîìåð ñëåäóþùåé ñòðîêè äëÿ âñòàâêè
                int nextRow = worksheet.Dimension.End.Row + 1;

                // Ïîëó÷èòå äîñòóï ê äàííûì èç DataGridView
                DataGridViewRow newDataGridViewRow = new DataGridViewRow();
                newDataGridViewRow.CreateCells(dataGridView1);

                // Äîáàâüòå âàøó ëîãèêó çäåñü äëÿ çàïîëíåíèÿ íîâîé ñòðîêè äàííûìè èç DataGridView
                // Íàïðèìåð, âû ìîæåòå ïðîéòè ïî ÿ÷åéêàì DataGridView è çàïîëíèòü íîâóþ ñòðîêó äàííûìè

                // Ïîñëå òîãî, êàê âû çàïîëíèòå íîâóþ ñòðîêó äàííûìè, âñòàâüòå å¸ â Excel
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    worksheet.Cells[nextRow, i + 1].Value = newDataGridViewRow.Cells[i].Value;
                }

                // Ñîõðàíèòå èçìåíåíèÿ â ôàéëå Excel
                excelPackage.Save();

                MessageBox.Show("Íîâàÿ ñòðîêà äîáàâëåíà â ôàéë Excel.");
            }
            else
            {
                MessageBox.Show("Ôàéë Excel íå çàãðóæåí.");
            }

        }
    }
}
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            try
            {
                // Ïîëó÷àåì âûáðàííûå ÿ÷åéêè
                DataGridViewSelectedCellCollection selectedCells = dataGridView1.SelectedCells;

                if (selectedCells.Count > 0)
                {
                    // Ïîëó÷àåì èíäåêñû âûáðàííûõ ÿ÷ååê
                    int rowIndex = selectedCells[0].RowIndex;
                    int columnIndex = selectedCells[0].ColumnIndex;

                    // Ïîëó÷àåì DataTable è óäàëÿåì ÿ÷åéêó
                    DataTable table = (DataTable)dataGridView1.DataSource;
                    table.Rows[rowIndex][columnIndex] = DBNull.Value;

                    // Îáíîâëÿåì DataGridView
                    dataGridView1.Refresh();
                }
                else
                {
                    MessageBox.Show("Âûáåðèòå ÿ÷åéêó äëÿ óäàëåíèÿ.", "Ïðåäóïðåæäåíèå", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Îøèáêà ïðè óäàëåíèè ÿ÷åéêè: " + ex.Message, "Îøèáêà", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void ñîõðàíèòüToolStripMenuItem_Click(object sender, EventArgs e)
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

        private void îáíîâèòüToolStripMenuItem_Click(object sender, EventArgs e)
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
            // Ïðåîáðàçóåì çíà÷åíèå ñëàéäåðà â êîýôôèöèåíò ìàñøòàáèðîâàíèÿ (îò 0.1 äî 2.0, íàïðèìåð)
            float scaleValue = 0.1f + (float)materialSlider1.Value / 50.0f;

            // Ïðèìåíÿåì ìàñøòàáèðîâàíèå ê ôîðìå
            this.Scale(new SizeF(scaleValue, scaleValue));

            // Îáíîâëÿåì ìàñøòàáèðîâàíèå è ðàñïîëîæåíèå êîíòðîëîâ â ôîðìå
            foreach (Control control in this.Controls)
            {
                control.Left = (int)(control.Left * scaleValue);
                control.Top = (int)(control.Top * scaleValue);
                control.Width = (int)(control.Width * scaleValue);
                control.Height = (int)(control.Height * scaleValue);

                // Äîïîëíèòåëüíûå íàñòðîéêè äëÿ îïðåäåëåííûõ òèïîâ êîíòðîëîâ
                if (control is TextBox)
                {
                    TextBox textBox = (TextBox)control;
                    textBox.Font = new Font(textBox.Font.FontFamily, textBox.Font.Size * scaleValue);
                }
                // Äîáàâüòå äîïîëíèòåëüíûå íàñòðîéêè äëÿ äðóãèõ òèïîâ êîíòðîëîâ, åñëè íåîáõîäèìî
            }
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }
    }
}

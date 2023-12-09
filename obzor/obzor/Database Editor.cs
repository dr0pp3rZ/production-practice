using MaterialSkin;
using MaterialSkin.Controls;
using obzor.Methods;
using System.Data;

namespace obzor
{
    public partial class Form1 : MaterialForm
    {
        private readonly MaterialSkinManager materialSkinManager; // Менеджер MaterialSkin
        private readonly FileOperationMethods _fileOperations = new();
        private readonly ButtonMethods _buttonmethods = new();

        public Form1()
        {
            InitializeComponent();
            materialSkinManager = MaterialSkinManager.Instance;
            ThemeMethods.MaterialSkinInit(materialSkinManager, this);
        }

        // Обработчик события изменения MaterialSwitch для изменения темы формы
        private void MaterialSwitch1_CheckedChanged(object sender, EventArgs e)
        {
            ThemeMethods.MaterialSwitchCheckedChanged(materialSwitch1, dataGridView1, materialSkinManager);
        }

        // Обработчик выбора листа из выпадающего списка
        private void ToolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            _fileOperations.ToolStripComboBoxSelectedIndexChanged(dataGridView1, toolStripComboBox1);
        }

        // Кнопка "Загрузить"
        private void ОткрытьToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            _buttonmethods.LoadFile(dataGridView1, openFileDialog1, toolStripComboBox1, toolStripMenuEditor);
        }

        // Кнопка "Сохранить"
        private void СохранитьToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            _buttonmethods.SaveButton(dataGridView1);
        }

        // Кнопка "Сохранить как"
        private void СохранитьКакToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1.DataSource != null && dataGridView1.DataSource is DataTable table)
            {
                DataTable selectedTable = table;
                _buttonmethods.SaveAsNewFile(dataGridView1, selectedTable, toolStripComboBox1, toolStripMenuEditor);
            }
            else
            {
                MessageBox.Show("Не удалось получить данные для создания файла.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Кнопка "Добавить"
        private void ДобавитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ButtonMethods.AddedButton(dataGridView1);
        }

        // Кнопка "Редактировать"
        private void РедактироватьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _buttonmethods.EditButton(dataGridView1, toolStripComboBox1, файлToolStripMenuItem, редактироватьToolStripMenuItem);
        }

        // Кнопка "Удалить"
        private void УдалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string fileName = string.Empty; // Переменная для хранения пути к файлу
            _buttonmethods.ClearSelectedCell(dataGridView1, fileName);
        }

        // Кнопка "Обновить"
        private void ОбновToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _buttonmethods.UpdateButton(dataGridView1);
        }

        // Обработка ошибок
        private void DataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            FileOperationMethods.DataError(dataGridView1, e);
        }
    }
}
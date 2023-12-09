using MaterialSkin;
using MaterialSkin.Controls;
using ColorScheme = MaterialSkin.ColorScheme;

namespace obzor.Methods
{
    internal class ThemeMethods
    {
        // Инициализация MaterialSkinManager
        public static void MaterialSkinInit(MaterialSkinManager materialSkinManager, MaterialForm form)
        {
            materialSkinManager.AddFormToManage(form);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(
                Primary.Orange800, Primary.Orange900,
                Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE
                );
        }

        // Обработчик события изменения MaterialSwitch для изменения темы формы
        public static void MaterialSwitchCheckedChanged(MaterialSwitch materialSwitch1, DataGridView dataGridView1, MaterialSkinManager materialSkinManager)
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

            UpdateDataGridViewColors(dataGridView1);
        }

        // Метод обновления цвета ячеек в DataGridView
        private static void UpdateDataGridViewColors(DataGridView dataGridView1)
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
    }
}

using System.Data;

namespace obzor.Methods
{
    internal class Variables
    {
        public DataTableCollection? tableCollection { get; set; } = null; // Коллекция таблиц из баз данных
        public DataTable? originalDataTable { get; set; } = null; // Исходная таблица для хранения оригинальных данных
        public string fileName { get; set; } = string.Empty; // Переменная для хранения пути к файлу
    }
}
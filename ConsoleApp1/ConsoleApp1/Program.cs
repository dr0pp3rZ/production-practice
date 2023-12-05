using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using ADOX;

namespace ExcelToAccessConverter
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            DisplayLoadingAnimation();

            do
            {
                // Разрешить пользователю выбрать файл Excel
                string excelFilePath = GetExcelFilePath();

                if (excelFilePath == null)
                {
                    Console.WriteLine("Отменено пользователем.");
                    return;
                }

                // Извлечение имени файла Excel без расширения
                string excelFileName = Path.GetFileNameWithoutExtension(excelFilePath);

                // Создание пути к файлу базы данных Access на основе имени файла Excel
                string accessFilePath = $"БАЗА\\{excelFileName}.accdb";
                CreateAccessDatabase(accessFilePath);

                // Создание подключений
                using (OleDbConnection excelConnection = new OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={excelFilePath};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1;\""))
                using (OleDbConnection accessConnection = new OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={accessFilePath}"))
                {
                    excelConnection.Open();
                    accessConnection.Open();

                    // Получение листов из файла Excel
                    DataTable excelSheets = excelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                    // Обработка каждого листа
                    foreach (DataRow row in excelSheets.Rows)
                    {
                        string sheetName = row["TABLE_NAME"].ToString();

                        // Удаление знака доллара и кавычек из имени листа
                        string cleanSheetName = sheetName.Replace("$", "").Replace("'", "");

                        // Позволяет пользователю указать пользовательское имя таблицы
                        Console.Write($"Введите имя таблицы для '{cleanSheetName}': ");
                        string customTableName = Console.ReadLine();

                        // Удаление таблицы, если она существует
                        DropTableIfExists(accessConnection, customTableName);

                        // Создание таблицы с пользовательским именем в файле Access (без знака доллара)
                        CreateTable(accessConnection, customTableName);

                        // Чтение данных из Excel и вставка их в Access
                        ReadDataFromExcelAndInsertIntoAccess(excelConnection, accessConnection, sheetName, customTableName);
                    }
                }

                Console.WriteLine("Конвертация завершена!");

                // Проверка, нужно ли конвертировать еще файлы
                Console.Write("Хотите конвертировать ещё файлы? (Y/N): ");
            } while (Console.ReadLine()?.ToUpper() == "Y");
        }

        private static void CreateAccessDatabase(string filePath)
        {
            Catalog catalog = new Catalog();

            try
            {
                // Создание новой базы данных Access
                if (!File.Exists(filePath))
                {
                    // Путь к базе данных включает подпапку "БАЗА"
                    string databaseFolderPath = Path.GetDirectoryName(filePath);

                    // Проверка существования папки "БАЗА"
                    if (!Directory.Exists(databaseFolderPath))
                    {
                        // Создание папки, если она не существует
                        Directory.CreateDirectory(databaseFolderPath);
                    }

                    catalog.Create($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Jet OLEDB:Engine Type=5");
                }
            }
            finally
            {
                // Освобождение объекта COM
                System.Runtime.InteropServices.Marshal.ReleaseComObject(catalog);
            }
        }

        private static void DropTableIfExists(OleDbConnection connection, string tableName)
        {
            try
            {
                // Проверка существования таблицы
                using (OleDbCommand checkTableCommand = new OleDbCommand($"SELECT COUNT(*) FROM [{tableName}]", connection))
                {
                    checkTableCommand.ExecuteNonQuery();
                }

                Console.WriteLine($"Таблица с именем '{tableName}' уже существует. Вы уверены, что хотите её перезаписать? (Y/N)");

                string response = Console.ReadLine();

                if (response?.ToUpper() != "Y")
                {
                    Console.WriteLine("Отказано в перезаписи таблицы, программа будет закрыта.");
                    Console.ReadKey();
                    Environment.Exit(0); // Выйти из программы, так как пользователь отказался перезаписывать
                }

                // Удаление таблицы, если она существует
                using (OleDbCommand dropTableCommand = new OleDbCommand($"DROP TABLE {tableName}", connection))
                {
                    dropTableCommand.ExecuteNonQuery();
                }
            }
            catch (OleDbException ex)
            {
                // Если возникает исключение, таблица не существует, и можно продолжить
            }
        }

        private static string GetAccessFilePath()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Title = "Выберите место сохранения базы данных Access",
                Filter = "Access Database (*.accdb)|*.accdb",
                FileName = "Новая_база_данных.accdb"
            };

            // Установим текущую директорию на рабочую директорию пользователя
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            DialogResult result = saveFileDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                string filePath = saveFileDialog.FileName;

                if (File.Exists(filePath))
                {
                    Console.WriteLine("Файл с таким именем уже существует. Что вы хотите сделать?");
                    Console.WriteLine("1. Изменить имя");
                    Console.WriteLine("2. Выбрать другой файл Excel");
                    Console.WriteLine("3. Выйти из программы");
                    Console.Write("Выберите опцию (1/2/3): ");

                    string choice = Console.ReadLine();

                    switch (choice)
                    {
                        case "1":
                            Console.WriteLine(); // Переход на новую строку для читаемости
                            return GetAccessFilePath(); // Запросить новое имя файла у пользователя
                        case "2":
                            Console.WriteLine(); // Переход на новую строку для читаемости
                            return GetExcelFilePath(); // Запросить новый файл Excel у пользователя
                        case "3":
                            Console.WriteLine(); // Переход на новую строку для читаемости
                            Console.WriteLine("Выход из программы.");
                            Environment.Exit(0); // Выйти из программы
                            break;
                        default:
                            Console.WriteLine("Некорректный выбор. Выход из программы.");
                            Environment.Exit(0);
                            break;
                    }
                }

                CreateAccessDatabase(filePath);  // Создайте базу данных после выбора места
                return filePath;
            }

            return null;
        }







        private static void CreateTable(OleDbConnection connection, string tableName)
        {
            using (OleDbCommand createTableCommand = new OleDbCommand($"CREATE TABLE [{tableName}] " +
                "(НОМЕР_ЗАЯВКИ int, [НОМЕР СОСТОЯНИЯ] int, [ДАТА СОЗДАНИЯ] datetime)", connection))
            {
                createTableCommand.ExecuteNonQuery();
            }
        }

        private static void ReadDataFromExcelAndInsertIntoAccess(OleDbConnection excelConnection, OleDbConnection accessConnection, string sheetName, string tableName)
        {
            using (OleDbCommand selectCommand = new OleDbCommand($"SELECT * FROM [{sheetName}]", excelConnection))
            using (OleDbDataReader reader = selectCommand.ExecuteReader())
            using (OleDbCommand insertCommand = new OleDbCommand($"INSERT INTO [{tableName}] (НОМЕР_ЗАЯВКИ, [НОМЕР СОСТОЯНИЯ], [ДАТА СОЗДАНИЯ]) VALUES (?, ?, ?)", accessConnection))
            {
                while (reader.Read())
                {
                    insertCommand.Parameters.Clear();
                    insertCommand.Parameters.AddWithValue("НОМЕР_ЗАЯВКИ", reader["НОМЕР_ЗАЯВКИ"]);
                    insertCommand.Parameters.AddWithValue("НОМЕР СОСТОЯНИЯ", reader["НОМЕР СОСТОЯНИЯ"]);
                    insertCommand.Parameters.AddWithValue("ДАТА СОЗДАНИЯ", reader["ДАТА СОЗДАНИЯ"]);
                    insertCommand.ExecuteNonQuery();
                }
            }
        }

        private static string GetExcelFilePath()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "Выберите файл Excel",
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                Multiselect = false
            };

            DialogResult result = openFileDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                return openFileDialog.FileName;
            }

            return null;
        }

        static void DisplayLoadingAnimation()
        {
            string[] lines = new string[]
            {
                
"██╗░░██╗███████╗██╗░░░░░██╗░░░░░███████╗██╗██████╗░███████╗  ░█████╗░██╗░░░░░██╗░░░██╗██████╗░",
"██║░░██║██╔════╝██║░░░░░██║░░░░░██╔════╝██║██╔══██╗██╔════╝  ██╔══██╗██║░░░░░██║░░░██║██╔══██╗",
"███████║█████╗░░██║░░░░░██║░░░░░█████╗░░██║██████╔╝█████╗░░  ██║░░╚═╝██║░░░░░██║░░░██║██████╦╝",
"██╔══██║██╔══╝░░██║░░░░░██║░░░░░██╔══╝░░██║██╔══██╗██╔══╝░░  ██║░░██╗██║░░░░░██║░░░██║██╔══██╗",
"██║░░██║███████╗███████╗███████╗██║░░░░░██║██║░░██║███████╗  ╚█████╔╝███████╗╚██████╔╝██████╦╝ ",
"╚═╝░░╚═╝╚══════╝╚══════╝╚══════╝╚═╝░░░░░╚═╝╚═╝░░╚═╝╚══════╝  ░╚════╝░╚══════╝░╚═════╝░╚═════╝░"
            };

            int delay = 1; // Задержка между появлением символов (в миллисекундах)

            Console.CursorVisible = false;

            // Рассчитываем координаты для вывода текста по центру консоли
            int top = Console.WindowHeight / 2 - lines.Length / 2;

            foreach (string line in lines)
            {
                int left = Console.WindowWidth / 2 - line.Length / 2;

                for (int i = 0; i < line.Length; i++)
                {
                    Console.SetCursorPosition(left + i, top);
                    Console.Write(line[i]);
                    Thread.Sleep(delay); // Задержка между появлением символов
                }

                top++;
            }

            // Загрузка
            string loadingText = "Loading...";
            int loadingLeft = Console.WindowWidth / 2 - loadingText.Length / 2;
            int loadingTop = Console.WindowHeight / 2 + lines.Length / 2 + 2;

            Console.SetCursorPosition(loadingLeft, loadingTop);
            Console.Write(loadingText);

            // Имитация загрузки
            for (int i = 0; i < 10; i++)
            {
                Console.SetCursorPosition(loadingLeft + loadingText.Length + i, loadingTop);
                Console.Write(".");
                Thread.Sleep(300); // Задержка между точками загрузки
            }

            Console.Clear(); // Очистить консоль после анимации
        }
    }

}
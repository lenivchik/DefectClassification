using System;
using System.IO;
using System.Linq;
using DefectClassification.Core;
using OfficeOpenXml;

namespace DefectClassification.TubeProcessor
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;

            // Set EPPlus license context
                       ExcelPackage.License.SetNonCommercialPersonal("Amir");

            Console.WriteLine("╔═══════════════════════════════════════════════════════════╗");
            Console.WriteLine("║  Обработка данных по трубкам                             ║");
            Console.WriteLine("║  Классификация дефектов с расчетом ширины                ║");
            Console.WriteLine("╚═══════════════════════════════════════════════════════════╝\n");

            string filePath;
            if (args.Length > 0)
            {
                filePath = args[0];
            }
            else
            {
                Console.Write("Введите путь к Excel файлу: ");
                filePath = Console.ReadLine()?.Trim() ?? "";
            }

            if (string.IsNullOrWhiteSpace(filePath))
            {
                Console.WriteLine("❌ Не указан путь к файлу");
                return;
            }

            ProcessTubeFile(filePath);

            Console.WriteLine("\nНажмите Enter для выхода...");
            Console.ReadLine();
        }

        static void ProcessTubeFile(string filePath)
        {
            try
            {
                // Validate file exists
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"❌ Файл не найден: {filePath}");
                    return;
                }

                // Validate extension
                var extension = Path.GetExtension(filePath).ToLower();
                if (extension != ".xlsx")
                {
                    Console.WriteLine("❌ Поддерживаются только .xlsx файлы");
                    return;
                }

                Console.WriteLine($"\n📂 Открытие файла: {Path.GetFileName(filePath)}");

                // Create backup
                var backupPath = CreateBackup(filePath);
                Console.WriteLine($"💾 Создана резервная копия: {Path.GetFileName(backupPath)}");

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var classifier = new DefectClassifier();
                    int totalProcessed = 0;
                    int totalErrors = 0;
                    int sheetsProcessed = 0;

                    // Dictionary to store defects found for each tube
                    var tubeDefects = new Dictionary<int, HashSet<string>>();

                    Console.WriteLine($"\n📊 Найдено листов: {package.Workbook.Worksheets.Count}");

                    foreach (var worksheet in package.Workbook.Worksheets)
                    {
                        var sheetName = worksheet.Name;

                        // Skip non-tube sheets
                        if (!sheetName.ToLower().Contains("трубка"))
                        {
                            Console.WriteLine($"⏩ Пропуск листа: {sheetName} (не содержит 'трубка')");
                            continue;
                        }

                        // Extract tube number from sheet name
                        var tubeNumber = ExtractTubeNumber(sheetName);
                        if (tubeNumber == -1)
                        {
                            Console.WriteLine($"⚠️  Не удалось извлечь номер трубки из: {sheetName}");
                            continue;
                        }

                        Console.WriteLine($"\n{'=' * 60}");
                        Console.WriteLine($"🔧 Обработка: {sheetName} (Трубка #{tubeNumber})");
                        Console.WriteLine($"{'=' * 60}");

                        // Find column headers
                        var config = FindTubeColumns(worksheet);
                        if (!config.IsValid)
                        {
                            Console.WriteLine($"⚠️  Не найдены столбцы 'Длина' и 'Площадь' - пропуск");
                            continue;
                        }

                        Console.WriteLine($"✓ Найдены столбцы:");
                        Console.WriteLine($"  - Длина: колонка {config.LengthColumn} (строка {config.HeaderRow})");
                        Console.WriteLine($"  - Площадь: колонка {config.AreaColumn} (строка {config.HeaderRow})");
                        Console.WriteLine($"  - Примечание: колонка {config.TextColumn} (строка {config.HeaderRow})");

                        // Process data rows
                        int startRow = config.HeaderRow + 1;
                        int rowCount = worksheet.Dimension?.Rows ?? 0;
                        int sheetProcessed = 0;
                        int sheetErrors = 0;

                        // Initialize defect set for this tube
                        if (!tubeDefects.ContainsKey(tubeNumber))
                        {
                            tubeDefects[tubeNumber] = new HashSet<string>();
                        }

                        for (int row = startRow; row <= rowCount; row++)
                        {
                            var lengthCell = worksheet.Cells[row, config.LengthColumn];
                            var areaCell = worksheet.Cells[row, config.AreaColumn];
                            var descCell = worksheet.Cells[row, config.TextColumn];

                            try
                            {
                                // Skip empty rows
                                if (string.IsNullOrWhiteSpace(lengthCell.Text) &&
                                    string.IsNullOrWhiteSpace(areaCell.Text))
                                {
                                    continue;
                                }

                                // Parse length (in mm)
                                if (!double.TryParse(lengthCell.Text, out double lengthMm) || lengthMm <= 0)
                                {
                                    descCell.Value = "ОШИБКА - Неверная длина";
                                    sheetErrors++;
                                    continue;
                                }

                                // Parse area (in sq.mm)
                                if (!double.TryParse(areaCell.Text, out double areaSqMm) || areaSqMm < 0)
                                {
                                    descCell.Value = "ОШИБКА - Неверная площадь";
                                    sheetErrors++;
                                    continue;
                                }

                                // Calculate width: Area / Length
                                double widthMm = areaSqMm / lengthMm;

                                // Convert to Lambda units (assuming 1 Lambda = 1mm for now)
                                double lengthLambda = lengthMm;
                                double widthLambda = widthMm;

                                // Classify defect
                                var region = classifier.Classify(lengthLambda, widthLambda, );
                                var description = DefectClassifier.GetRegionDescription(region);

                                // Write results
                                descCell.Value = description;

                                // Add to defect set (skip "Нет деффектов")
                                if (!description.ToLower().Contains("ошибка"))
                                {
                                    tubeDefects[tubeNumber].Add(description);
                                }

                                sheetProcessed++;
                            }
                            catch (Exception ex)
                            {
                                descCell.Value = $"ОШИБКА: {ex.Message}";
                                sheetErrors++;
                            }
                        }

                        Console.WriteLine($"✓ Обработано строк: {sheetProcessed}");
                        if (sheetErrors > 0)
                        {
                            Console.WriteLine($"  ⚠️  Ошибок: {sheetErrors}");
                        }

                        // Show statistics for this sheet
                        ShowSheetStatistics(worksheet, config, sheetProcessed, config.TextColumn);

                        // Show found defects for this tube
                        if (tubeDefects[tubeNumber].Any())
                        {
                            Console.WriteLine($"  🔍 Найденные дефекты: {string.Join(", ", tubeDefects[tubeNumber])}");
                        }
                        else
                        {
                            Console.WriteLine($"  ✓ Дефектов не обнаружено");
                        }

                        totalProcessed += sheetProcessed;
                        totalErrors += sheetErrors;
                        sheetsProcessed++;
                    }

                    // Update ИНТЕРВАЛЫ sheet
                    UpdateIntervalsSheet(package, tubeDefects);

                    // Save the file
                    package.Save();

                    Console.WriteLine($"\n{'=' * 60}");
                    Console.WriteLine("✅ ОБРАБОТКА ЗАВЕРШЕНА!");
                    Console.WriteLine($"{'=' * 60}");
                    Console.WriteLine($"  Обработано листов: {sheetsProcessed}");
                    Console.WriteLine($"  Всего строк: {totalProcessed}");
                    if (totalErrors > 0)
                    {
                        Console.WriteLine($"  ⚠️  Всего ошибок: {totalErrors}");
                    }
                    Console.WriteLine($"  Результаты сохранены в: {filePath}");
                    Console.WriteLine($"  Резервная копия: {backupPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n❌ ОШИБКА: {ex.Message}");
                Console.WriteLine($"   {ex.GetType().Name}");

                if (ex.InnerException != null)
                {
                    Console.WriteLine($"   Внутренняя ошибка: {ex.InnerException.Message}");
                }
            }
        }

        static int ExtractTubeNumber(string sheetName)
        {
            // Extract number from sheet name like "144 трубка" or "1 трубка"
            var parts = sheetName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var part in parts)
            {
                if (int.TryParse(part, out int number))
                {
                    return number;
                }
            }
            return -1;
        }

        static void UpdateIntervalsSheet(ExcelPackage package, Dictionary<int, HashSet<string>> tubeDefects)
        {
            // Find the ИНТЕРВАЛЫ sheet (case-insensitive)
            var intervalsSheet = package.Workbook.Worksheets
                .FirstOrDefault(ws => ws.Name.ToUpper().Contains("ИНТЕРВАЛ"));

            if (intervalsSheet == null)
            {
                Console.WriteLine("\n⚠️  Лист 'ИНТЕРВАЛЫ' не найден - пропуск агрегации");
                return;
            }

            Console.WriteLine($"\n{'=' * 60}");
            Console.WriteLine($"📝 Обновление листа: {intervalsSheet.Name}");
            Console.WriteLine($"{'=' * 60}");

            // Find the columns for tube number and notes
            int tubeNumCol = -1;
            int notesCol = -1;
            int headerRow = -1;

            // Search for headers in first 15 rows
            for (int row = 1; row <= Math.Min(15, intervalsSheet.Dimension?.Rows ?? 0); row++)
            {
                for (int col = 1; col <= (intervalsSheet.Dimension?.Columns ?? 0); col++)
                {
                    var cellText = intervalsSheet.Cells[row, col].Text?.Trim().ToLower() ?? "";

                    if (cellText.Contains("трубки") && cellText.Contains("№"))
                    {
                        tubeNumCol = col;
                        headerRow = row;
                    }
                    else if (cellText.Contains("примечание"))
                    {
                        notesCol = col;
                        if (headerRow == 0)
                            headerRow = row;
                    }
                }
            }

            if (tubeNumCol == -1 || notesCol == -1 || headerRow == -1)
            {
                Console.WriteLine("⚠️  Не найдены столбцы '№ трубки' и 'Примечание' - пропуск");
                return;
            }

            Console.WriteLine($"✓ Найдены столбцы:");
            Console.WriteLine($"  - № трубки: колонка {tubeNumCol}");
            Console.WriteLine($"  - Примечание: колонка {notesCol}");
            Console.WriteLine($"  - Заголовки в строке: {headerRow}");

            int updatedCount = 0;
            int startRow = headerRow + 2; // Skip header and subheader rows
            int maxRow = intervalsSheet.Dimension?.Rows ?? 0;

            for (int row = startRow; row <= maxRow; row++)
            {
                var tubeNumCell = intervalsSheet.Cells[row, tubeNumCol];
                var tubeNumText = tubeNumCell.Text?.Trim() ?? "";

                // Skip empty rows
                if (string.IsNullOrWhiteSpace(tubeNumText))
                    continue;

                // Try to parse tube number
                if (int.TryParse(tubeNumText, out int tubeNumber))
                {
                    if (tubeDefects.ContainsKey(tubeNumber) && tubeDefects[tubeNumber].Any())
                    {
                        var defectsList = string.Join(", ", tubeDefects[tubeNumber].OrderBy(d => d));
                        intervalsSheet.Cells[row, notesCol].Value = defectsList;
                        updatedCount++;
                        Console.WriteLine($"  ✓ Трубка {tubeNumber}: {defectsList}");
                    }
                }
            }

            Console.WriteLine($"\n✅ Обновлено записей в листе ИНТЕРВАЛЫ: {updatedCount}");
        }

        static TubeColumnConfiguration FindTubeColumns(ExcelWorksheet worksheet)
        {
            var config = new TubeColumnConfiguration();

            // Search in first 10 rows for headers
            for (int row = 1; row <= Math.Min(10, worksheet.Dimension?.Rows ?? 0); row++)
            {
                for (int col = 1; col <= (worksheet.Dimension?.Columns ?? 0); col++)
                {
                    var cellText = worksheet.Cells[row, col].Text?.Trim().ToLower() ?? "";

                    if (cellText.Contains("длина"))
                    {
                        config.LengthColumn = col;
                        config.HeaderRow = row;
                    }
                    else if (cellText.Contains("площадь") || cellText.Contains("пло-\nщадь"))
                    {
                        config.AreaColumn = col;
                        if (config.HeaderRow == 0)
                            config.HeaderRow = row;
                    }
                    else if (cellText.Contains("примеч.") || cellText.Contains("примечание"))
                    {
                        config.TextColumn = col;
                        if (config.HeaderRow == 0)
                            config.HeaderRow = row;
                    }
                }
            }

            return config;
        }

        static void ShowSheetStatistics(ExcelWorksheet worksheet, TubeColumnConfiguration config,
                                       int totalRows, int typeCol)
        {
            var statistics = new Dictionary<string, int>();
            int startRow = config.HeaderRow + 1;

            for (int row = startRow; row <= worksheet.Dimension?.Rows; row++)
            {
                var typeText = worksheet.Cells[row, typeCol].Text;
                if (!string.IsNullOrWhiteSpace(typeText) && !typeText.Contains("ОШИБКА"))
                {
                    if (!statistics.ContainsKey(typeText))
                    {
                        statistics[typeText] = 0;
                    }
                    statistics[typeText]++;
                }
            }

            if (statistics.Any())
            {
                Console.WriteLine("\n  Статистика по типам:");
                foreach (var kvp in statistics.OrderByDescending(x => x.Value))
                {
                    var percentage = totalRows > 0 ? (double)kvp.Value / totalRows * 100.0 : 0;
                    Console.WriteLine($"    {kvp.Key,-25} : {kvp.Value,4} ({percentage,5:F1}%)");
                }
            }
        }

        static string CreateBackup(string filePath)
        {
            var directory = Path.GetDirectoryName(filePath) ?? "";
            var fileName = Path.GetFileNameWithoutExtension(filePath);
            var extension = Path.GetExtension(filePath);
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            var backupPath = Path.Combine(directory, $"{fileName}_backup_{timestamp}{extension}");
            File.Copy(filePath, backupPath, true);

            return backupPath;
        }
    }

    class TubeColumnConfiguration
    {
        public int LengthColumn { get; set; }
        public int AreaColumn { get; set; }
        public int TextColumn { get; set; }
        public int HeaderRow { get; set; }

        public bool IsValid => LengthColumn > 0 && AreaColumn > 0 && HeaderRow > 0;
    }
}
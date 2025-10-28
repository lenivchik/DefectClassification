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

                        Console.WriteLine(new string('═', 59) + "\n");
                        Console.WriteLine($"🔧 Обработка: {sheetName}");
                        Console.WriteLine(new string('═', 59) + "\n");

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
                        Console.WriteLine($"  - Примеч: колонка {config.TextColumn} (строка {config.HeaderRow})");
                        Console.WriteLine($"  - Потеря: колонка {config.MaxMetLoss} (строка {config.HeaderRow})");


                        // Add result columns if needed
                        //int typeCol = config.AreaColumn + 1;
                        //int descCol = config.AreaColumn + 2;
                        //int widthCol = config.AreaColumn + 3;

                        //if (string.IsNullOrWhiteSpace(worksheet.Cells[config.HeaderRow, typeCol].Text))
                        //{
                        //    worksheet.Cells[config.HeaderRow, typeCol].Value = "Тип\nдефекта";
                        //    worksheet.Cells[config.HeaderRow, typeCol].Style.Font.Bold = true;
                        //    worksheet.Cells[config.HeaderRow, typeCol].Style.WrapText = true;
                        //}
                        //if (string.IsNullOrWhiteSpace(worksheet.Cells[config.HeaderRow, descCol].Text))
                        //{
                        //    worksheet.Cells[config.HeaderRow, descCol].Value = "Описание";
                        //    worksheet.Cells[config.HeaderRow, descCol].Style.Font.Bold = true;
                        //}
                        //if (string.IsNullOrWhiteSpace(worksheet.Cells[config.HeaderRow, widthCol].Text))
                        //{
                        //    worksheet.Cells[config.HeaderRow, widthCol].Value = "Ширина\n(выч.)";
                        //    worksheet.Cells[config.HeaderRow, widthCol].Style.Font.Bold = true;
                        //    worksheet.Cells[config.HeaderRow, widthCol].Style.WrapText = true;
                        //}

                        // Process data rows
                        int startRow = config.HeaderRow + 1;
                        int rowCount = worksheet.Dimension?.Rows ?? 0;
                        int sheetProcessed = 0;
                        int sheetErrors = 0;

                        for (int row = startRow; row <= rowCount; row++)
                        {
                            var lengthCell = worksheet.Cells[row, config.LengthColumn];
                            var areaCell = worksheet.Cells[row, config.AreaColumn];
                            var descCell = worksheet.Cells[row, config.TextColumn];
                            var maxMetCell = worksheet.Cells[row, config.MaxMetLoss];
                            try
                            {

                                // Skip empty rows
                                if (string.IsNullOrWhiteSpace(lengthCell.Text) &&
                                    string.IsNullOrWhiteSpace(areaCell.Text) &&
                                    string.IsNullOrWhiteSpace(descCell.Text) &&
                                    string.IsNullOrWhiteSpace(maxMetCell.Text))
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

                                if (!double.TryParse(maxMetCell.Text, out double maxLoss) || maxLoss < 0 || maxLoss > 100)
                                {
                                    descCell.Value = "ОШИБКА - Неверная потеря";
                                    sheetErrors++;
                                    continue;
                                }
                                // Calculate width: Area / Length
                                double widthMm = areaSqMm / lengthMm;

                                // Convert to Lambda units (assuming 1 Lambda = 1mm for now)
                                // You may need to adjust this conversion factor
                                double lengthLambda = lengthMm / 10;
                                double widthLambda = widthMm / 10;

                                // Classify defect
                                var region = classifier.Classify(lengthLambda, widthLambda, maxLoss);
                                var description = DefectClassifier.GetRegionDescription(region);

                                // Write results
                                descCell.Value = description;

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

                        totalProcessed += sheetProcessed;
                        totalErrors += sheetErrors;
                        sheetsProcessed++;

                        // Auto-fit columns
                        worksheet.Column(config.TextColumn).AutoFit();


                    }

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
                    //Console.WriteLine($"  Резервная копия: {backupPath}");
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
                    else if (cellText.Contains("потеря"))
                    {
                        config.MaxMetLoss = col;
                        if (config.HeaderRow == 0)
                            config.HeaderRow = row;
                    }
                }
            }

            return config;
        }

        static void ApplyColorCoding(ExcelRange cell, DefectRegion region)
        {
            switch (region)
            {
                case DefectRegion.ExtСor:
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 200, 200));
                    break;
                case DefectRegion.PointСor:
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(200, 255, 200));
                    break;
                case DefectRegion.LongSlit:
                case DefectRegion.TranSlit:
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 255, 200));
                    break;
                case DefectRegion.LongGroov:
                case DefectRegion.TranGroov:
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(200, 220, 255));
                    break;
                case DefectRegion.Ulcer:
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 255));
                    break;
            }
        }

        static void ShowSheetStatistics(ExcelWorksheet worksheet, TubeColumnConfiguration config,
                                       int totalRows, int typeCol)
        {
            var statistics = new System.Collections.Generic.Dictionary<string, int>();
            int startRow = config.HeaderRow + 1;

            for (int row = startRow; row <= worksheet.Dimension?.Rows; row++)
            {
                var typeText = worksheet.Cells[row, typeCol].Text;
                if (!string.IsNullOrWhiteSpace(typeText) && typeText != "ОШИБКА")
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
                    Console.WriteLine($"    {kvp.Key,-12} : {kvp.Value,4} ({percentage,5:F1}%)");
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
        public int MaxMetLoss { get; set; }
        public int TextColumn { get; set; }
        public int HeaderRow { get; set; }

        public bool IsValid => LengthColumn > 0 && AreaColumn > 0 && HeaderRow > 0;
    }
}
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

                        Console.WriteLine(new string('═', 59) + "\n");
                        Console.WriteLine($"🔧 Обработка: {sheetName}");
                        Console.WriteLine(new string('═', 59) + "\n");

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
                        Console.WriteLine($"  - Примеч: колонка {config.TextColumn} (строка {config.HeaderRow})");
                        Console.WriteLine($"  - Потеря: колонка {config.MaxMetLoss} (строка {config.HeaderRow})");


                        // Add result columns if needed
                        //int typeCol = config.AreaColumn + 1;
                        //int descCol = config.AreaColumn + 2;
                        //int widthCol = config.AreaColumn + 3;

                        //if (string.IsNullOrWhiteSpace(worksheet.Cells[config.HeaderRow, typeCol].Text))
                        //{
                        //    worksheet.Cells[config.HeaderRow, typeCol].Value = "Тип\nдефекта";
                        //    worksheet.Cells[config.HeaderRow, typeCol].Style.Font.Bold = true;
                        //    worksheet.Cells[config.HeaderRow, typeCol].Style.WrapText = true;
                        //}
                        //if (string.IsNullOrWhiteSpace(worksheet.Cells[config.HeaderRow, descCol].Text))
                        //{
                        //    worksheet.Cells[config.HeaderRow, descCol].Value = "Описание";
                        //    worksheet.Cells[config.HeaderRow, descCol].Style.Font.Bold = true;
                        //}
                        //if (string.IsNullOrWhiteSpace(worksheet.Cells[config.HeaderRow, widthCol].Text))
                        //{
                        //    worksheet.Cells[config.HeaderRow, widthCol].Value = "Ширина\n(выч.)";
                        //    worksheet.Cells[config.HeaderRow, widthCol].Style.Font.Bold = true;
                        //    worksheet.Cells[config.HeaderRow, widthCol].Style.WrapText = true;
                        //}

                        // Process data rows
                        int startRow = config.HeaderRow + 1;
                        int rowCount = worksheet.Dimension?.Rows ?? 0;
                        int sheetProcessed = 0;
                        int sheetErrors = 0;

                        for (int row = startRow; row <= rowCount; row++)
                        {
                            var lengthCell = worksheet.Cells[row, config.LengthColumn];
                            var areaCell = worksheet.Cells[row, config.AreaColumn];
                            var descCell = worksheet.Cells[row, config.TextColumn];
                            var maxMetCell = worksheet.Cells[row, config.MaxMetLoss];
                            try
                            {

                                // Skip empty rows
                                if (string.IsNullOrWhiteSpace(lengthCell.Text) &&
                                    string.IsNullOrWhiteSpace(areaCell.Text) &&
                                    string.IsNullOrWhiteSpace(descCell.Text) &&
                                    string.IsNullOrWhiteSpace(maxMetCell.Text))
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

                                if (!double.TryParse(maxMetCell.Text, out double maxLoss) || maxLoss < 0 || maxLoss > 100)
                                {
                                    descCell.Value = "ОШИБКА - Неверная потеря";
                                    sheetErrors++;
                                    continue;
                                }
                                // Calculate width: Area / Length
                                double widthMm = areaSqMm / lengthMm;

                                // Convert to Lambda units (assuming 1 Lambda = 1mm for now)
                                // You may need to adjust this conversion factor
                                double lengthLambda = lengthMm / 10;
                                double widthLambda = widthMm / 10;

                                // Classify defect
                                var region = classifier.Classify(lengthLambda, widthLambda, maxLoss);
                                var description = DefectClassifier.GetRegionDescription(region);

                                // Write results
                                descCell.Value = description;

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

                        totalProcessed += sheetProcessed;
                        totalErrors += sheetErrors;
                        sheetsProcessed++;

                        // Auto-fit columns
                        worksheet.Column(config.TextColumn).AutoFit();


                    }

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
                    //Console.WriteLine($"  Резервная копия: {backupPath}");
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
                    else if (cellText.Contains("потеря"))
                    {
                        config.MaxMetLoss = col;
                        if (config.HeaderRow == 0)
                            config.HeaderRow = row;
                    }
                }
            }

            return config;
        }

        static void ApplyColorCoding(ExcelRange cell, DefectRegion region)
        {
            switch (region)
            {
                case DefectRegion.ExtСor:
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 200, 200));
                    break;
                case DefectRegion.PointСor:
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(200, 255, 200));
                    break;
                case DefectRegion.LongSlit:
                case DefectRegion.TranSlit:
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 255, 200));
                    break;
                case DefectRegion.LongGroov:
                case DefectRegion.TranGroov:
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(200, 220, 255));
                    break;
                case DefectRegion.Ulcer:
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 220, 255));
                    break;
            }
        }

        static void ShowSheetStatistics(ExcelWorksheet worksheet, TubeColumnConfiguration config,
                                       int totalRows, int typeCol)
        {
            var statistics = new System.Collections.Generic.Dictionary<string, int>();
            int startRow = config.HeaderRow + 1;

            for (int row = startRow; row <= worksheet.Dimension?.Rows; row++)
            {
                var typeText = worksheet.Cells[row, typeCol].Text;
                if (!string.IsNullOrWhiteSpace(typeText) && typeText != "ОШИБКА")
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
                    Console.WriteLine($"    {kvp.Key,-12} : {kvp.Value,4} ({percentage,5:F1}%)");
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
        public int MaxMetLoss { get; set; }
        public int TextColumn { get; set; }
        public int HeaderRow { get; set; }

        public bool IsValid => LengthColumn > 0 && AreaColumn > 0 && HeaderRow > 0;
    }
}
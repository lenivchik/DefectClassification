using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DefectClassification.Core;
using OfficeOpenXml;

namespace DefectClassification.GUI.Services
{
    public class ExcelProcessor
    {
        public void ProcessFile(
            string filePath,
            double lambdaThreshold,
            bool isStaticMode,
            bool createBackup,
            Dictionary<int, double>? tubeThicknesses = null,
            IProgress<(int current, int total, string message)>? progress = null)
        {
            // Set EPPlus license
            ExcelPackage.License.SetNonCommercialPersonal("Amir");

            // Validate file
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Файл не найден: {filePath}");

            var extension = Path.GetExtension(filePath).ToLower();
            if (extension != ".xlsx")
                throw new ArgumentException("Поддерживаются только .xlsx файлы");

            // Create backup if requested
            if (createBackup)
            {
                var backupPath = CreateBackup(filePath);
                progress?.Report((0, 100, $"Создана резервная копия: {Path.GetFileName(backupPath)}"));
            }

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var classifier = new DefectClassifier();
                var tubeDefects = new Dictionary<int, Dictionary<string, List<(double depth, double maxLoss)>>>();

                // Count tube sheets
                var tubeSheets = package.Workbook.Worksheets
                    .Where(ws => ws.Name.ToLower().Contains("трубка"))
                    .ToList();

                int totalSheets = tubeSheets.Count;
                int processedSheets = 0;

                foreach (var worksheet in tubeSheets)
                {
                    var sheetName = worksheet.Name;
                    var tubeNumber = ExtractTubeNumber(sheetName);

                    if (tubeNumber == -1)
                    {
                        progress?.Report((processedSheets, totalSheets,
                            $"⚠️ Не удалось извлечь номер трубки из: {sheetName}"));
                        continue;
                    }

                    // Get wall thickness for this tube
                    double wallThickness;
                    if (isStaticMode)
                    {
                        // Use static lambda threshold
                        wallThickness = lambdaThreshold;
                    }
                    else
                    {
                        // Try to get from dynamic configuration
                        if (tubeThicknesses != null && tubeThicknesses.ContainsKey(tubeNumber))
                        {
                            wallThickness = tubeThicknesses[tubeNumber];
                        }
                        else
                        {
                            progress?.Report((processedSheets, totalSheets,
                                $"⚠️ Трубка #{tubeNumber} не найдена в конфигурации - пропуск"));
                            processedSheets++;
                            continue;
                        }
                    }

                    progress?.Report((processedSheets, totalSheets,
                        $"Обработка: {sheetName} (Трубка #{tubeNumber}, толщина {wallThickness} мм)"));

                    // Find columns
                    var config = FindTubeColumns(worksheet);
                    if (!config.IsValid)
                    {
                        progress?.Report((processedSheets, totalSheets,
                            $"⚠️ Не найдены столбцы в листе {sheetName}"));
                        processedSheets++;
                        continue;
                    }

                    // Initialize defect dictionary for this tube
                    if (!tubeDefects.ContainsKey(tubeNumber))
                    {
                        tubeDefects[tubeNumber] = new Dictionary<string, List<(double depth, double maxLoss)>>();
                    }

                    // Process rows
                    int startRow = config.HeaderRow + 1;
                    int rowCount = worksheet.Dimension?.Rows ?? 0;

                    for (int row = startRow; row <= rowCount; row++)
                    {
                        var depthCell = worksheet.Cells[row, config.DepthColumn];
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

                            // Parse max loss
                            if (!double.TryParse(maxMetCell.Text, out double maxLoss) ||
                                maxLoss < 0 || maxLoss > 100)
                            {
                                descCell.Value = "ОШИБКА - Неверная потеря";
                                continue;
                            }

                            if (maxLoss < 40)
                            {
                                continue;
                            }

                            // Parse depth (in meters)
                            if (!double.TryParse(depthCell.Text, out double depthM))
                            {
                                depthM = 0;
                            }

                            // Parse length (in mm)
                            if (!double.TryParse(lengthCell.Text, out double lengthMm) || lengthMm <= 0)
                            {
                                descCell.Value = "ОШИБКА - Неверная длина";
                                continue;
                            }

                            // Parse area (in sq.mm)
                            if (!double.TryParse(areaCell.Text, out double areaSqMm) || areaSqMm < 0)
                            {
                                descCell.Value = "ОШИБКА - Неверная площадь";
                                continue;
                            }

                            // Calculate width: Area / Length
                            double widthMm = areaSqMm / lengthMm;

                            // Convert to Lambda units using wall thickness
                            // Lambda = measurement / wall_thickness
                            double lengthLambda = lengthMm / wallThickness;
                            double widthLambda = widthMm / wallThickness;

                            // Classify defect
                            var region = classifier.Classify(lengthLambda, widthLambda);
                            var description = DefectClassifier.GetRegionDescription(region);

                            // Write results
                            descCell.Value = description;

                            // Add to defect dictionary
                            if (!description.ToLower().Contains("нет деффектов") &&
                                !description.ToLower().Contains("ошибка"))
                            {
                                if (!tubeDefects[tubeNumber].ContainsKey(description))
                                {
                                    tubeDefects[tubeNumber][description] = new List<(double depth, double maxLoss)>();
                                }
                                tubeDefects[tubeNumber][description].Add((depthM, maxLoss));
                            }
                        }
                        catch (Exception ex)
                        {
                            descCell.Value = $"ОШИБКА: {ex.Message}";
                        }
                    }

                    // Auto-fit columns
                    worksheet.Column(config.TextColumn).AutoFit();

                    processedSheets++;
                    progress?.Report((processedSheets, totalSheets,
                        $"Обработано: {sheetName} ({processedSheets}/{totalSheets})"));
                }

                // Update ИНТЕРВАЛЫ sheet
                progress?.Report((totalSheets, totalSheets, "Обновление листа ИНТЕРВАЛЫ..."));
                UpdateIntervalsSheet(package, tubeDefects);

                // Save the file
                progress?.Report((totalSheets, totalSheets, "Сохранение файла..."));
                package.Save();
            }
        }

        private int ExtractTubeNumber(string sheetName)
        {
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

        private void UpdateIntervalsSheet(
            ExcelPackage package,
            Dictionary<int, Dictionary<string, List<(double depth, double maxLoss)>>> tubeDefects)
        {
            var intervalsSheet = package.Workbook.Worksheets
                .FirstOrDefault(ws => ws.Name.ToUpper().Contains("ИНТЕРВАЛ"));

            if (intervalsSheet == null)
                return;

            // Find columns
            int tubeNumCol = -1;
            int notesCol = -1;
            int headerRow = -1;

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
                    else if (cellText.Contains("примечание") || cellText.Contains("примеч"))
                    {
                        notesCol = col;
                        if (headerRow == 0)
                            headerRow = row;
                    }
                }
            }

            if (tubeNumCol == -1 || notesCol == -1 || headerRow == -1)
                return;

            int startRow = headerRow + 2;
            int maxRow = intervalsSheet.Dimension?.Rows ?? 0;

            for (int row = startRow; row <= maxRow; row++)
            {
                var tubeNumCell = intervalsSheet.Cells[row, tubeNumCol];
                var tubeNumText = tubeNumCell.Text?.Trim() ?? "";

                if (string.IsNullOrWhiteSpace(tubeNumText))
                    continue;

                if (int.TryParse(tubeNumText, out int tubeNumber))
                {
                    if (tubeDefects.ContainsKey(tubeNumber) && tubeDefects[tubeNumber].Any())
                    {
                        var defectList = new List<(string defect, double depth, double maxLoss)>();
                        foreach (var kvp in tubeDefects[tubeNumber])
                        {
                            foreach (var item in kvp.Value)
                            {
                                defectList.Add((kvp.Key, item.depth, item.maxLoss));
                            }
                        }

                        var sortedDefects = defectList.OrderBy(d => d.depth)
                            .Select(d => $"\"{d.defect}\" (макс потеря {d.maxLoss}%, на глубине {d.depth})")
                            .ToList();
                        var defectsList = string.Join(", ", sortedDefects);

                        intervalsSheet.Cells[row, notesCol].Value = defectsList;
                    }
                }
            }

            intervalsSheet.Column(notesCol).Width = 100;

            for (int row = startRow; row <= maxRow; row++)
            {
                intervalsSheet.Cells[row, notesCol].Style.WrapText = true;
            }
        }

        private TubeColumnConfiguration FindTubeColumns(ExcelWorksheet worksheet)
        {
            var config = new TubeColumnConfiguration();

            for (int row = 1; row <= Math.Min(10, worksheet.Dimension?.Rows ?? 0); row++)
            {
                for (int col = 1; col <= (worksheet.Dimension?.Columns ?? 0); col++)
                {
                    var cellText = worksheet.Cells[row, col].Text?.Trim().ToLower() ?? "";

                    if (cellText.Contains("глубина"))
                    {
                        config.DepthColumn = col;
                        if (config.HeaderRow == 0)
                            config.HeaderRow = row;
                    }
                    else if (cellText.Contains("длина"))
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

        private string CreateBackup(string filePath)
        {
            var directory = Path.GetDirectoryName(filePath) ?? "";
            var fileName = Path.GetFileNameWithoutExtension(filePath);
            var extension = Path.GetExtension(filePath);
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            var backupPath = Path.Combine(directory, $"{fileName}_backup_{timestamp}{extension}");
            File.Copy(filePath, backupPath, true);

            return backupPath;
        }

        private class TubeColumnConfiguration
        {
            public int LengthColumn { get; set; }
            public int AreaColumn { get; set; }
            public int DepthColumn { get; set; }
            public int MaxMetLoss { get; set; }
            public int TextColumn { get; set; }
            public int HeaderRow { get; set; }

            public bool IsValid => LengthColumn > 0 && AreaColumn > 0 && HeaderRow > 0;
        }
    }
}
using DefectClassification.Core;
using DefectClassification.Core.Exceptions;
using System;
using System.Linq;
using DefectClassification.WellSim;




namespace DefectClassification.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            System.Console.OutputEncoding = System.Text.Encoding.UTF8;

            System.Console.WriteLine("╔═══════════════════════════════════════════════════════════╗");
            System.Console.WriteLine("║  Система классификации дефектов обсадной колонны         ║");
            System.Console.WriteLine("║  С симуляцией скважины (1000м, шаг 2.5мм)                ║");
            System.Console.WriteLine("╚═══════════════════════════════════════════════════════════╝\n");

            while (true)
            {
                System.Console.WriteLine("\n═══ ГЛАВНОЕ МЕНЮ ═══");
                System.Console.WriteLine("1. Классификация отдельного дефекта");
                System.Console.WriteLine("2. Симуляция скважины с случайными дефектами");
                System.Console.WriteLine("3. Добавить дефекты вручную в скважину");
                System.Console.WriteLine("4. Анализ дефектов в диапазоне глубин");
                System.Console.WriteLine("5. Экспорт дефектов в CSV");
                System.Console.WriteLine("0. Выход");
                System.Console.Write("\nВыберите опцию: ");

                var choice = System.Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        ClassifySingleDefect();
                        break;
                    case "2":
                        SimulateWellWithRandomDefects();
                        break;
                    case "3":
                        ManuallyAddDefects();
                        break;
                    case "4":
                        AnalyzeDepthRange();
                        break;
                    case "5":
                        ExportDefectsToCSV();
                        break;
                    case "0":
                        System.Console.WriteLine("\nДо свидания!");
                        return;
                    default:
                        System.Console.WriteLine("❌ Неверный выбор. Попробуйте снова.");
                        break;
                }
            }
        }

        static void ClassifySingleDefect()
        {
            var classifier = new DefectClassifier();

            System.Console.WriteLine("\n--- Классификация отдельного дефекта ---");

            try
            {
                System.Console.Write("Введите ДЛИНУ дефекта (0-10 Λ): ");
                if (!double.TryParse(System.Console.ReadLine(), out double length))
                {
                    System.Console.WriteLine("❌ Неверный формат числа.");
                    return;
                }

                System.Console.Write("Введите ШИРИНУ дефекта (0-10 Λ): ");
                if (!double.TryParse(System.Console.ReadLine(), out double width))
                {
                    System.Console.WriteLine("❌ Неверный формат числа.");
                    return;
                }

                var region = classifier.Classify(length, width);
                var description = DefectClassifier.GetRegionDescription(region);

                System.Console.WriteLine($"\n✓ Результат:");
                System.Console.WriteLine($"  Длина: {length:F2} Λ, Ширина: {width:F2} Λ");
                System.Console.WriteLine($"  Тип: {region}");
                System.Console.WriteLine($"  Описание: {description}");
            }
            catch (InvalidDefectMeasurementException ex)
            {
                System.Console.WriteLine($"❌ Ошибка: {ex.Message}");
            }
        }

        static void SimulateWellWithRandomDefects()
        {
            System.Console.WriteLine("\n--- Симуляция скважины ---");
            System.Console.Write("Введите количество дефектов для генерации: ");

            if (!int.TryParse(System.Console.ReadLine(), out int defectCount) || defectCount <= 0)
            {
                System.Console.WriteLine("❌ Неверное количество дефектов.");
                return;
            }

            System.Console.WriteLine($"\n⏳ Создание скважины (1000м, шаг 2.5мм = 400,000 точек)...");
            var well = new Well();

            System.Console.WriteLine($"⏳ Генерация {defectCount} случайных дефектов...");
            well.GenerateRandomDefects(defectCount, seed: (int)DateTime.Now.Ticks);

            System.Console.WriteLine("✓ Генерация завершена!\n");

            var stats = well.GetStatistics();
            System.Console.WriteLine(stats.ToString());

            System.Console.WriteLine("\n--- Примеры дефектов ---");
            var sampleDefects = well.GetAllDefects().Take(10).ToList();
            foreach (var defect in sampleDefects)
            {
                var description = DefectClassifier.GetRegionDescription(defect.DefectType!.Value);
                System.Console.WriteLine($"Глубина {defect.DepthMeters:F2}м: {description} " +
                                       $"(Д={defect.DefectLength:F2}, Ш={defect.DefectWidth:F2})");
            }

            if (well.GetAllDefects().Count() > 10)
            {
                System.Console.WriteLine($"... и еще {well.GetAllDefects().Count() - 10} дефектов");
            }
        }

        static void ManuallyAddDefects()
        {
            System.Console.WriteLine("\n--- Добавление дефектов вручную ---");
            var well = new Well();

            while (true)
            {
                System.Console.Write("\nВведите глубину в метрах (или 'q' для завершения): ");
                var depthInput = System.Console.ReadLine();

                if (depthInput?.ToLower() == "q")
                    break;

                if (!double.TryParse(depthInput, out double depth) || depth < 0 || depth > 1000)
                {
                    System.Console.WriteLine("❌ Неверная глубина (должна быть 0-1000м).");
                    continue;
                }

                System.Console.Write("Длина дефекта (Λ): ");
                if (!double.TryParse(System.Console.ReadLine(), out double length))
                {
                    System.Console.WriteLine("❌ Неверный формат.");
                    continue;
                }

                System.Console.Write("Ширина дефекта (Λ): ");
                if (!double.TryParse(System.Console.ReadLine(), out double width))
                {
                    System.Console.WriteLine("❌ Неверный формат.");
                    continue;
                }

                try
                {
                    well.AddDefect(depth, length, width);
                    var measurement = well.GetMeasurementAtDepth(depth);
                    var description = DefectClassifier.GetRegionDescription(measurement.DefectType!.Value);
                    System.Console.WriteLine($"✓ Добавлен дефект: {description}");
                }
                catch (Exception ex)
                {
                    System.Console.WriteLine($"❌ Ошибка: {ex.Message}");
                }
            }

            if (well.GetAllDefects().Any())
            {
                System.Console.WriteLine("\n" + well.GetStatistics().ToString());
            }
        }

        static void AnalyzeDepthRange()
        {
            System.Console.WriteLine("\n--- Анализ диапазона глубин ---");
            var well = new Well();

            System.Console.Write("Сколько случайных дефектов сгенерировать? ");
            if (!int.TryParse(System.Console.ReadLine(), out int count) || count <= 0)
            {
                System.Console.WriteLine("❌ Неверное количество.");
                return;
            }

            well.GenerateRandomDefects(count);

            System.Console.Write("\nНачальная глубина (м): ");
            if (!double.TryParse(System.Console.ReadLine(), out double startDepth))
            {
                System.Console.WriteLine("❌ Неверный формат.");
                return;
            }

            System.Console.Write("Конечная глубина (м): ");
            if (!double.TryParse(System.Console.ReadLine(), out double endDepth))
            {
                System.Console.WriteLine("❌ Неверный формат.");
                return;
            }

            var measurementsInRange = well.GetMeasurementsInRange(startDepth, endDepth).ToList();
            var defectsInRange = measurementsInRange.Where(m => m.HasDefect).ToList();

            System.Console.WriteLine($"\n✓ Диапазон {startDepth:F2}м - {endDepth:F2}м:");
            System.Console.WriteLine($"  Всего точек измерения: {measurementsInRange.Count}");
            System.Console.WriteLine($"  Дефектов найдено: {defectsInRange.Count}");

            if (defectsInRange.Any())
            {
                System.Console.WriteLine("\n--- Дефекты в диапазоне ---");
                foreach (var defect in defectsInRange.Take(20))
                {
                    var description = DefectClassifier.GetRegionDescription(defect.DefectType!.Value);
                    System.Console.WriteLine($"  {defect.DepthMeters:F2}м: {description}");
                }

                if (defectsInRange.Count > 20)
                {
                    System.Console.WriteLine($"  ... и еще {defectsInRange.Count - 20} дефектов");
                }
            }
        }

        static void ExportDefectsToCSV()
        {
            System.Console.WriteLine("\n--- Экспорт дефектов в CSV ---");
            var well = new Well();

            System.Console.Write("Количество случайных дефектов: ");
            if (!int.TryParse(System.Console.ReadLine(), out int count) || count <= 0)
            {
                System.Console.WriteLine("❌ Неверное количество.");
                return;
            }

            well.GenerateRandomDefects(count);

            var csv = well.ExportDefectsToCSV();
            var filename = $"well_defects_{DateTime.Now:yyyyMMdd_HHmmss}.csv";

            try
            {
                System.IO.File.WriteAllText(filename, csv);
                System.Console.WriteLine($"✓ Экспортировано {count} дефектов в файл: {filename}");
            }
            catch (Exception ex)
            {
                System.Console.WriteLine($"❌ Ошибка при сохранении: {ex.Message}");
            }
        }
    }
}
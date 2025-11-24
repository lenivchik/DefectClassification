using DefectClassification.Core.Exceptions;

namespace DefectClassification.Core
{
    /// <summary>
    /// Классификатор.
    /// </summary>
    public class DefectClassifier
    {
        private const double MinMeasurement = 0.0;
        private const double MaxMeasurement = 10.0;

        /// <summary>
        /// Классификация на основе длины/ширины.
        /// </summary>
        /// <param name="length"> Длина - ось X.</param>
        /// <param name="width"> Ширина - ось Y.</param>
        /// <param name="thickness"> Толщина стенки.</param>

        /// <returns>Определенние дефекта.</returns>
        public DefectRegion Classify(double length, double width, double MaxMet)
        {
            
            ValidateMeasurement(length, nameof(length));
            ValidateMeasurement(width, nameof(width));
            

            // 1: Обширная Коррозия (Длина/Ширина ≥ 3)
            if (length >=3.0 && width >= 3.0)
                return DefectRegion.ExtСor;

            // 2: Точечная Коррозия (Длина/Ширина ≤ 1)
            else if (length <= 1.0 && width <= 1.0)
                return DefectRegion.PointСor;

            // 3: Продольный Шлиц ( Длина > 1; Ширина ≤ 1)
            else if (length > 1.0 && width <= 1.0)
                return DefectRegion.LongSlit;

            // 4: Поперечный Шлиц ( Длина ≤ 1; Ширина > 1)
            else if (length <= 1.0 && width > 1.0)
                return DefectRegion.TranSlit;

            // 5: Продольная канавка (Длина * 0.5 ≤ Ширина)
            else if (length * 0.5 >= width)
                return DefectRegion.LongGroov;

            // 6: Поперечная канавка (Длина * 2 ≥ Ширина)
            else if (length * 2 <= width)
                return DefectRegion.TranGroov;

            // Остальное: Язва
            return DefectRegion.Ulcer;
        }
        public DefectRegion Classify(double length, double width)
        {

            ValidateMeasurement(length, nameof(length));
            ValidateMeasurement(width, nameof(width));

            // 1: Обширная Коррозия (Длина/Ширина ≥ 3)
            if (length >= 3.0 && width >= 3.0)
                return DefectRegion.ExtСor;

            // 2: Точечная Коррозия (Длина/Ширина ≤ 1)
            else if (length <= 1.0 && width <= 1.0)
                return DefectRegion.PointСor;

            // 3: Продольный Шлиц ( Длина > 1; Ширина ≤ 1)
            else if (length > 1.0 && width <= 1.0)
                return DefectRegion.LongSlit;

            // 4: Поперечный Шлиц ( Длина ≤ 1; Ширина > 1)
            else if (length <= 1.0 && width > 1.0)
                return DefectRegion.TranSlit;

            // 5: Продольная канавка (Длина * 0.5 ≤ Ширина)
            else if (length * 0.5 >= width)
                return DefectRegion.LongGroov;

            // 6: Поперечная канавка (Длина * 2 ≥ Ширина)
            else if (length * 2 <= width)
                return DefectRegion.TranGroov;

            // Остальное: Язва
            return DefectRegion.Ulcer;
        }
        private void ValidateMeasurement(double value, string parameterName)
        {
            if (value <= MinMeasurement)
                throw new InvalidDefectMeasurementException(parameterName, value);

            if (double.IsNaN(value) || double.IsInfinity(value))
                throw new InvalidDefectMeasurementException(
                    $"Неверная {parameterName}: {value}. Должно быть число.",
                    parameterName, value);
        }

        /// <summary>
        /// Вывод в "читаемом" варианте.
        /// </summary>
        public static string GetRegionDescription(DefectRegion region) => region switch
        {
            DefectRegion.Clear => "Нет деффектов",
            DefectRegion.ExtСor => "Обширная коррозия",
            DefectRegion.PointСor=> "Точечная коррозия",
            DefectRegion.LongSlit => "Продольный шлиц",
            DefectRegion.TranSlit => "Поперечный шлиц",
            DefectRegion.LongGroov => "Продольная канавка",
            DefectRegion.TranGroov => "Поперечная канавка",
            DefectRegion.Ulcer => "Язва",
            _ => "Unknown"
        };
    }
}
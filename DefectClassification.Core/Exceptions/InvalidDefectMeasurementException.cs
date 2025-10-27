namespace DefectClassification.Core.Exceptions
{
    /// <summary>
    /// Ошибки в параметрах.
    /// </summary>
    public class InvalidDefectMeasurementException : ArgumentException
    {
        public string MeasurementName { get; }
        public double InvalidValue { get; }

        public InvalidDefectMeasurementException(string measurementName, double value)
            : base($"Неверная {measurementName}: {value}. Должно быть число.", measurementName)
        {
            MeasurementName = measurementName;
            InvalidValue = value;
        }

        public InvalidDefectMeasurementException(string message, string measurementName, double value)
            : base(message, measurementName)
        {
            MeasurementName = measurementName;
            InvalidValue = value;
        }
    }
}
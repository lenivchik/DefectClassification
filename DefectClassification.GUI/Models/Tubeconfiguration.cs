using CommunityToolkit.Mvvm.ComponentModel;

namespace DefectClassification.GUI.Models
{
    /// <summary>
    /// Конфигурация трубки с номером и толщиной стенки
    /// </summary>
    public partial class TubeConfiguration : ObservableObject
    {
        [ObservableProperty]
        private int _tubeNumber;

        [ObservableProperty]
        private double _wallThickness = 10.0;

        public TubeConfiguration()
        {
        }

        public TubeConfiguration(int tubeNumber, double wallThickness = 10.0)
        {
            TubeNumber = tubeNumber;
            WallThickness = wallThickness;
        }
    }
}
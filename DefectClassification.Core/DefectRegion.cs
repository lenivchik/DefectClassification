using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DefectClassification.Core
{
    /// <summary>
    /// Классификации дефектов обсадной колонны по размерам (оси X, Y)
    /// </summary>
    public enum DefectRegion
    {
        /// <summary>
        /// Обширная коррозия (Extensive corrosion) - Длина/Ширина ≥ 3
        /// </summary>
        ExtСor,

        /// <summary>
        /// Точечная коррозия (Point corrosion) - Длина/Ширина ≤ 1
        /// </summary>
        PointСor,

        /// <summary>
        /// Продольный шлиц (Longitudinal slit) - Длина > 1; Ширина ≤ 1.
        /// </summary>
        LongSlit,

        /// <summary>
        /// Поперечный шлиц (Transverse slit) - Длина ≤ 1; Ширина > 1.
        /// </summary>
        TranSlit,

        /// <summary>
        ///  Продольная канавка (Longitudinal groove) - Длина * 0.5 ≤ Ширина.
        /// </summary>
        LongGroov,

        /// <summary>
        ///  Поперечная канавка (Transverse groove) - Длина * 2 ≥ Ширина.
        /// </summary>
        TranGroov,

        /// <summary>
        /// Язва (Ulcer) - Остальные случаи.
        /// </summary>
        Ulcer
    }
}


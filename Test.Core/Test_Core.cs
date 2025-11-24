using DefectClassification.Core;
using DefectClassification.Core.Exceptions;
using Xunit;

namespace Test
{
    public class Test_Core
    {
        private readonly DefectClassifier _classifier = new();

        // Обширная коррозия
        [Fact]
        public void Classify_ExtensiveCorrosion_AtBoundary_ReturnsExtensiveCorrosion()
        {
            var result = _classifier.Classify(3.0, 3.0);
            Assert.Equal(DefectRegion.ExtСor, result);
        }

        [Fact]
        public void Classify_ExtensiveCorrosion_AboveBoundary_ReturnsExtensiveCorrosion()
        {
            var result = _classifier.Classify(5.0, 5.0);
            Assert.Equal(DefectRegion.ExtСor, result);
        }

        [Fact]
        public void Classify_ExtensiveCorrosion_LargeWidth_ReturnsExtensiveCorrosion()
        {
            var result = _classifier.Classify(3.0, 10.0);
            Assert.Equal(DefectRegion.ExtСor, result);
        }

        [Fact]
        public void Classify_ExtensiveCorrosion_LargeLength_ReturnsExtensiveCorrosion()
        {
            var result = _classifier.Classify(10.0, 3.0);
            Assert.Equal(DefectRegion.ExtСor, result);
        }


        // Точечная коррозия 
        [Fact]
        public void Classify_PointCorrosion_SmallDimensions_ReturnsPointCorrosion()
        {
            var result = _classifier.Classify(1.0, 1.0);
            Assert.Equal(DefectRegion.PointСor, result);
        }

        [Fact]
        public void Classify_PointCorrosion_AtOrigin_ReturnsPointCorrosion()
        {
            var result = _classifier.Classify(0.5, 0.5);
            Assert.Equal(DefectRegion.PointСor, result);
        }

        [Fact]
        public void Classify_PointCorrosion_NearBoundary_ReturnsPointCorrosion()
        {
            var result = _classifier.Classify(0.5, 1.0);
            Assert.Equal(DefectRegion.PointСor, result);
        }

        // LONGITUDINAL SLIT TESTS (4 tests)
        [Fact]
        public void Classify_LongitudinalSlit_LongNarrow_ReturnsLongitudinalSlit()
        {
            var result = _classifier.Classify(5.0, 0.8);
            Assert.Equal(DefectRegion.LongSlit, result);
        }

        [Fact]
        public void Classify_LongitudinalSlit_AtLengthBoundary_ReturnsLongitudinalSlit()
        {
            var result = _classifier.Classify(4.0, 1.0);
            Assert.Equal(DefectRegion.LongSlit, result);
        }

        [Fact]
        public void Classify_LongitudinalSlit_MaxLength_ReturnsLongitudinalSlit()
        {
            var result = _classifier.Classify(8.0, 0.1);
            Assert.Equal(DefectRegion.LongSlit, result);
        }

        [Fact]
        public void Classify_LongitudinalSlit_NearWidthBoundary_ReturnsLongitudinalSlit()
        {
            var result = _classifier.Classify(6.0, 1.4);
            Assert.Equal(DefectRegion.LongSlit, result);
        }

        // LONGITUDINAL GROOVE TESTS (4 tests)
        [Fact]
        public void Classify_LongitudinalGroove_ModerateWidth_ReturnsLongitudinalGroove()
        {
            var result = _classifier.Classify(3.0, 2.0);
            Assert.Equal(DefectRegion.LongGroov, result);
        }

        [Fact]
        public void Classify_LongitudinalGroove_AtLowerWidthBoundary_ReturnsLongitudinalGroove()
        {
            var result = _classifier.Classify(5.0, 1.5);
            Assert.Equal(DefectRegion.LongGroov, result);
        }

        [Fact]
        public void Classify_LongitudinalGroove_NearUpperBoundary_ReturnsLongitudinalGroove()
        {
            var result = _classifier.Classify(4.0, 3.9);
            Assert.Equal(DefectRegion.LongGroov, result);
        }

        [Fact]
        public void Classify_LongitudinalGroove_AtLengthBoundary_ReturnsLongitudinalGroove()
        {
            var result = _classifier.Classify(2.0, 2.5);
            Assert.Equal(DefectRegion.LongGroov, result);
        }

        // ULCER TESTS (3 tests)
        [Fact]
        public void Classify_Ulcer_MediumDimensions_ReturnsUlcer()
        {
            var result = _classifier.Classify(3.0, 1.0);
            Assert.Equal(DefectRegion.Ulcer, result);
        }

        [Fact]
        public void Classify_Ulcer_ShortButWide_ReturnsUlcer()
        {
            var result = _classifier.Classify(1.5, 2.5);
            Assert.Equal(DefectRegion.Ulcer, result);
        }

        [Fact]
        public void Classify_Ulcer_EdgeCase_ReturnsUlcer()
        {
            var result = _classifier.Classify(1.0, 3.5);
            Assert.Equal(DefectRegion.Ulcer, result);
        }

        // VALIDATION TESTS (5 tests)
        [Fact]
        public void Classify_NegativeLength_ThrowsException()
        {
            var ex = Assert.Throws<InvalidDefectMeasurementException>(
                () => _classifier.Classify(-1.0, 2.0));
            Assert.Equal("length", ex.MeasurementName);
            Assert.Equal(-1.0, ex.InvalidValue);
        }

        [Fact]
        public void Classify_NegativeWidth_ThrowsException()
        {
            var ex = Assert.Throws<InvalidDefectMeasurementException>(
                () => _classifier.Classify(2.0, -0.5));
            Assert.Equal("width", ex.MeasurementName);
        }

        [Fact]
        public void Classify_LengthTooLarge_ThrowsException()
        {
            var ex = Assert.Throws<InvalidDefectMeasurementException>(
                () => _classifier.Classify(9.0, 2.0));
            Assert.Equal("length", ex.MeasurementName);
            Assert.Equal(9.0, ex.InvalidValue);
        }

        [Fact]
        public void Classify_WidthTooLarge_ThrowsException()
        {
            Assert.Throws<InvalidDefectMeasurementException>(
                () => _classifier.Classify(2.0, 8.5));
        }

        [Fact]
        public void Classify_NaNValue_ThrowsException()
        {
            Assert.Throws<InvalidDefectMeasurementException>(
                () => _classifier.Classify(double.NaN, 2.0));
        }

        // EDGE CASES (2 tests)
        [Fact]
        public void Classify_MaxValues_ReturnsExtensiveCorrosion()
        {
            var result = _classifier.Classify(8.0, 8.0);
            Assert.Equal(DefectRegion.ExtСor, result);
        }

        //[Fact]
        //public void GetRegionDescription_AllRegions_ReturnsDescriptions()
        //{
        //    Assert.Contains("Extensive", DefectClassifier.GetRegionDescription(DefectRegion.ОбширнаяКоррозия));
        //    Assert.Contains("Point", DefectClassifier.GetRegionDescription(DefectRegion.ТочечнаяКоррозия));
        //    Assert.Contains("Slit", DefectClassifier.GetRegionDescription(DefectRegion.ПродольныйШлиц));
        //    Assert.Contains("Groove", DefectClassifier.GetRegionDescription(DefectRegion.ПродольнаяКанавка));
        //    Assert.Contains("Ulcer", DefectClassifier.GetRegionDescription(DefectRegion.Язва));
        //}
    }
}
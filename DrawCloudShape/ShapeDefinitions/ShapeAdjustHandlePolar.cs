namespace EPPlusImageRenderer.ShapeDefinitions
{
    public class ShapeAdjustHandlePolar : ShapeAdjustHandleBase
    {
        public override ShapeAdjustHandleType AhType => ShapeAdjustHandleType.Polar;
        public string AngleAdjustmentGuide { get; set; }
        public string RadialAdjustmentGuide { get; set; }
        public object MaximumAngleAdjustment { get; set; }
        public object MaximumRadialAdjustment { get; set; }
        public object MinimumAngleAdjustment { get; set; }
        public object MinimumRadialAdjustment { get; set; }
        internal override ShapeAdjustHandleBase Clone()
        {
            return new ShapeAdjustHandlePolar()
            {
                AngleAdjustmentGuide = AngleAdjustmentGuide,
                RadialAdjustmentGuide = RadialAdjustmentGuide,
                MaximumAngleAdjustment = MaximumAngleAdjustment,
                MinimumAngleAdjustment = MinimumAngleAdjustment,
                MaximumRadialAdjustment = MaximumRadialAdjustment,
                MinimumRadialAdjustment = MinimumRadialAdjustment,
            };
        }

    }
}

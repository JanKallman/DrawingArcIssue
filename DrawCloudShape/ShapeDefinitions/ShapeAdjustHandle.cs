using System.Xml.Linq;
using System;
using System.Reflection.Metadata;
using System.Xml;

namespace EPPlusImageRenderer.ShapeDefinitions
{

    public abstract class ShapeAdjustHandleBase
    {
        public abstract ShapeAdjustHandleType AhType { get; }
        public ShapePositionCoordinate PositionCoordinate { get; set; }
        internal static ShapeAdjustHandleXY CreateXy(XmlReader xr)
        {
            return new ShapeAdjustHandleXY()
            {
                HorizontalAdjustmentGuide = xr.GetAttribute("gdRefX"),
                VerticalAdjustmentGuide = xr.GetAttribute("gdRefY"),
                MinimumHorizontalAdjustment = xr.GetAttribute("minX"),
                MaximumHorizontalAdjustment = xr.GetAttribute("minY"),
                MinimumVerticalAdjustment = xr.GetAttribute("maxX"),
                MaximumVerticalAdjustment = xr.GetAttribute("maxY"),
            };
        }
        internal static ShapeAdjustHandlePolar CreatePolar(XmlReader xr)
        {
            return new ShapeAdjustHandlePolar()
            {
                AngleAdjustmentGuide = xr.GetAttribute("gdRefAng"),
                RadialAdjustmentGuide = xr.GetAttribute("gdRefR"),
                MinimumAngleAdjustment = xr.GetAttribute("minAng"),
                MinimumRadialAdjustment = xr.GetAttribute("minR"),
                MaximumAngleAdjustment = xr.GetAttribute("maxAng"),
                MaximumRadialAdjustment = xr.GetAttribute("maxR"),
            };
        }

        internal abstract ShapeAdjustHandleBase Clone();
    }
}

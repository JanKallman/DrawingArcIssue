using OfficeOpenXml.Utils;
using System.Globalization;
using System.Xml;
namespace OfficeOpenXml.Drawing
{
    public enum PathDrawingType
    {
        MoveTo,
        LineTo,
        ArcTo,
        CubicBezTo,
        QuadBezerTo,
        Close
    }
    /// <summary>
    /// How a shape path is filled.
    /// </summary>
    public enum PathFillMode
    {
        /// <summary>
        /// The corresponding path should have a normally shaded color applied to it’s fill
        /// </summary>
        Norm,
        /// <summary>
        /// The corresponding path should have a darker shaded color applied to it’s fill.
        /// </summary>
        Darken,
        /// <summary>
        /// The corresponding path should have a slightly darker shaded color applied to it’s fill.
        /// </summary>
        DarkenLess,
        /// <summary>
        /// The corresponding path should have a lightly shaded color applied to it’s fill.
        /// </summary>
        Lighten,
        /// <summary>
        /// The corresponding path should have a slightly lighter shaded color applied to it’s fill.
        /// </summary>
        LightenLess,
        /// <summary>
        /// The corresponding path should have no fill.
        /// </summary>
        None
    }
    internal class DrawCoordinate
    {
        public DrawCoordinate(DrawCoordinate c)
        {
            X = c.X;
            Y = c.Y;
            XName = c.XName;
            YName = c.YName;
        }

        public DrawCoordinate(object x, object y)
        {
            if (x is long xl)
            {
                X = xl;
            }
            else
            {
                XName = x.ToString();
                X = null;
            }
            if (y is long yl)
            {
                Y = yl;
            }
            else
            {
                YName = y.ToString();
                Y = null;
            }

        }
        public double? X { get; set; }
        public double? Y { get; set; }
        public string XName { get; set; }
        public string YName { get; set; }
    }
    public abstract class PathsBase
    {
        public abstract PathDrawingType Type { get; }

        internal abstract PathsBase Clone();
        public abstract double EndX { get; }
        public abstract double EndY { get; }
    }
    internal abstract class PathWithCoordinates : PathsBase
    {
        protected PathWithCoordinates(XmlElement e)
        {
            foreach (var cn in e.ChildNodes)
            {
                if (cn is XmlElement ce && ce.LocalName == "pt")
                {
                    Coordinates.Add(new DrawCoordinate(GetNameOrNumber(ce.GetAttribute("x")), GetNameOrNumber(ce.GetAttribute("y"))));
                }
            }
        }

        private object GetNameOrNumber(string s)
        {
            if (long.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out var l))
            {
                return l;
            }
            return s;
        }

        protected PathWithCoordinates(XmlReader xr)
        {
            var name = xr.LocalName;
            while (xr.Read())
            {
                if (xr.LocalName == "pt" && xr.NodeType == XmlNodeType.Element)
                {
                    Coordinates.Add(new DrawCoordinate(GetNameOrNumber(xr.GetAttribute("x")), GetNameOrNumber(xr.GetAttribute("y"))));
                }
                else if (xr.IsEndElementWithName(name))
                {
                    break;
                }
            }
        }

        protected PathWithCoordinates(PathWithCoordinates clone)
        {
            foreach (var c in clone.Coordinates)
            {
                Coordinates.Add(new DrawCoordinate(c));
            }
        }
        public List<DrawCoordinate> Coordinates { get; set; } = new List<DrawCoordinate>();
        public override double EndX => Coordinates.Count > 0D ? Coordinates[Coordinates.Count - 1].X.Value : 0D;
        public override double EndY => Coordinates.Count > 0D ? Coordinates[Coordinates.Count - 1].Y.Value : 0D;
    }
    internal class MoveTo : PathWithCoordinates
    {
        public MoveTo(MoveTo clone) : base(clone)
        {

        }
        public MoveTo(XmlElement e) : base(e)
        {
        }
        public MoveTo(XmlReader xr) : base(xr)
        {
        }
        public override PathDrawingType Type => PathDrawingType.MoveTo;
        public DrawCoordinate Coordinate { get; set; }

        internal override PathsBase Clone()
        {
            return new MoveTo(this);
        }
    }
    internal class LineTo : PathWithCoordinates
    {
        public LineTo(LineTo clone) : base(clone)
        {

        }
        public LineTo(XmlReader xr) : base(xr)
        {

        }

        public LineTo(XmlElement e) : base(e)
        {

        }
        public override PathDrawingType Type => PathDrawingType.LineTo;
        public DrawCoordinate Coordinate { get; set; }
        internal override PathsBase Clone()
        {
            return new LineTo(this);
        }
    }
    internal class ClosePath : PathsBase
    {
        public ClosePath()
        {

        }
        public override PathDrawingType Type => PathDrawingType.Close;
        internal override PathsBase Clone()
        {
            return new ClosePath();
        }
        public override double EndX => double.MinValue;
        public override double EndY => double.MinValue;
    }
    internal class QuadBezerTo : PathWithCoordinates
    {
        public QuadBezerTo(QuadBezerTo clone) : base(clone)
        {

        }
        public QuadBezerTo(XmlReader xr) : base(xr)
        {

        }
        public QuadBezerTo(XmlElement e) : base(e)
        {

        }
        public override PathDrawingType Type => PathDrawingType.QuadBezerTo;
        internal override PathsBase Clone()
        {
            return new QuadBezerTo(this);
        }

    }

    internal class CubicBezerTo : PathWithCoordinates
    {
        public CubicBezerTo(CubicBezerTo clone) : base(clone)
        {

        }
        public CubicBezerTo(XmlReader xr) : base(xr)
        {

        }
        public CubicBezerTo(XmlElement e) : base(e)
        {

        }

        public override PathDrawingType Type => PathDrawingType.CubicBezTo;
        internal override PathsBase Clone()
        {
            return new CubicBezerTo(this);
        }
    }

    internal class ArcTo : PathsBase
    {
        public ArcTo(XmlReader xr)
        {
            if (long.TryParse(xr.GetAttribute("hR"), out var hrv))
            {
                HeightRadius = hrv;
            }
            else
            {
                HeightRadiusName = xr.GetAttribute("hR");
            }

            if (long.TryParse(xr.GetAttribute("wR"), out var wrv))
            {
                WidthRadius = wrv;
            }
            else
            {
                WidthRadiusName = xr.GetAttribute("wR");
            }

            if (long.TryParse(xr.GetAttribute("swAng"), out var swAng))
            {
                SwingAngle = swAng;
            }
            else
            {
                SwingAngleName = xr.GetAttribute("swAng");
            }

            if (long.TryParse(xr.GetAttribute("stAng"), out var stAng))
            {
                StartAngle = stAng;
            }
            else
            {
                StartAngleName = xr.GetAttribute("stAng");
            }
        }
        public ArcTo(XmlElement e)
        {
            if (long.TryParse(e.GetAttribute("hR"), out var hrv))
            {
                HeightRadius = hrv;
            }
            else
            {
                HeightRadiusName = e.GetAttribute("hR");
            }

            if (long.TryParse(e.GetAttribute("wR"), out var wrv))
            {
                WidthRadius = wrv;
            }
            else
            {
                WidthRadiusName = e.GetAttribute("wR");
            }

            if (long.TryParse(e.GetAttribute("swAng"), out var swAng))
            {
                SwingAngle = swAng;
            }
            else
            {
                SwingAngleName = e.GetAttribute("swAng");
            }

            if (long.TryParse(e.GetAttribute("stAng"), out var stAng))
            {
                StartAngle = stAng;
            }
            else
            {
                StartAngleName = e.GetAttribute("stAng");
            }
        }
        public override PathDrawingType Type => PathDrawingType.ArcTo;
        public double? HeightRadius { get; set; }
        public double? StartAngle { get; set; }
        public double? SwingAngle { get; set; }
        public double? WidthRadius { get; set; }
        public string HeightRadiusName { get; set; }
        public string StartAngleName { get; set; }
        public string SwingAngleName { get; set; }
        public string WidthRadiusName { get; set; }
        private ArcTo()
        {

        }
        internal override PathsBase Clone()
        {
            return new ArcTo()
            {
                HeightRadius = HeightRadius,
                StartAngle = StartAngle,
                SwingAngle = SwingAngle,
                WidthRadius = WidthRadius,
                HeightRadiusName = HeightRadiusName,
                StartAngleName = StartAngleName,
                SwingAngleName = SwingAngleName,
                WidthRadiusName = WidthRadiusName
            };
        }
        double _endX, _endY;
        internal void SetEndCoordinates(double x, double y)
        {
            _endX = x;
            _endY = y;
        }
        public override double EndX => _endX;
        public override double EndY => _endY;
    }
    internal class DrawingPath
    {
        public DrawingPath(DrawingPath clone)
        {
            Width = clone.Width;
            Height = clone.Height;
            Fill = clone.Fill;
            Stroke = clone.Stroke;
            ExtrusionOk = clone.ExtrusionOk;
            foreach (var p in clone.Paths)
            {
                Paths.Add(p.Clone());
            }
        }
        public DrawingPath(XmlReader xr)
        {
            Width = ConvertUtil.GetValueLongNull(xr.GetAttribute("w"));
            Height = ConvertUtil.GetValueLongNull(xr.GetAttribute("h"));
            Fill = GetFill(xr.GetAttribute("fill"));
            Stroke = ConvertUtil.ToBooleanString(xr.GetAttribute("stroke"), true);
            ExtrusionOk = ConvertUtil.ToBooleanString(xr.GetAttribute("extrusionOk"), false);
            while (xr.Read())
            {
                if (xr.NodeType == XmlNodeType.Element)
                {
                    switch (xr.LocalName)
                    {
                        case "moveTo":
                            Paths.Add(new MoveTo(xr));
                            break;
                        case "lnTo":
                            Paths.Add(new LineTo(xr));
                            break;
                        case "cubicBezTo":
                            Paths.Add(new CubicBezerTo(xr));
                            break;
                        case "quadBezTo":
                            Paths.Add(new QuadBezerTo(xr));
                            break;
                        case "arcTo":
                            Paths.Add(new ArcTo(xr));
                            break;
                        case "close":
                            Paths.Add(new ClosePath());
                            break;
                    }
                }
                else if (xr.LocalName == "path" && xr.NodeType == XmlNodeType.EndElement)
                {
                    break;
                }
            }
        }

        public DrawingPath(XmlElement topNode, XmlNamespaceManager nsm)
        {
            Width = int.Parse(topNode.GetAttribute("w"));
            Height = int.Parse(topNode.GetAttribute("h"));
            Fill = GetFill(topNode.GetAttribute("fill"));
            Stroke = ConvertUtil.ToBooleanString(topNode.GetAttribute("stroke"), true);
            ExtrusionOk = ConvertUtil.ToBooleanString(topNode.GetAttribute("extrusionOk"), true);
            foreach (var child in topNode.ChildNodes)
            {
                if (child is XmlElement e)
                {
                    switch (e.LocalName)
                    {
                        case "moveTo":
                            Paths.Add(new MoveTo(e));
                            break;
                        case "lnTo":
                            Paths.Add(new LineTo(e));
                            break;
                        case "cubicBezTo":
                            Paths.Add(new CubicBezerTo(e));
                            break;
                        case "quadBezTo":
                            Paths.Add(new CubicBezerTo(e));
                            break;
                        case "arcTo":
                            Paths.Add(new ArcTo(e));
                            break;
                        case "close":
                            Paths.Add(new ClosePath());
                            break;
                    }
                }
            }
        }

        private PathFillMode GetFill(string s)
        {
            if (string.IsNullOrEmpty(s) == false)
            {
                return (PathFillMode)Enum.Parse(typeof(PathFillMode), s, true);
            }
            return PathFillMode.Norm;
        }

        internal DrawingPath Clone() => new DrawingPath(this);

        public bool Stroke { get; set; }
        public bool ExtrusionOk { get; set; }
        public PathFillMode Fill { get; set; }
        public double? Width { get; set; }
        public double? Height { get; set; }
        public List<PathsBase> Paths { get; set; } = new List<PathsBase>();
    }
}
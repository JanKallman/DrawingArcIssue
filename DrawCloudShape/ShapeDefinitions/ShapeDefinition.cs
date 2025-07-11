using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System.Diagnostics;
using System.Globalization;

namespace EPPlusImageRenderer.ShapeDefinitions
{
    [DebuggerDisplay("{Style}")]
    internal class ShapeDefinition
    {
        internal Dictionary<string, double> _calculatedValues = new Dictionary<string, double>();

        internal ShapeDefinition()
        {
                
        }
        internal ShapeDefinition(ShapeDefinition clone)
        {
            Style=clone.Style;
            if (clone.ShapeAdjustValues != null)
            {
                ShapeAdjustValues = new List<ShapeGuide>();
                foreach (var av in clone.ShapeAdjustValues)
                {
                    ShapeAdjustValues.Add(av.Clone());
                }
            }
            if (clone.ShapeGuides != null)
            {
                ShapeGuides = new List<ShapeGuide>();
                foreach (var g in clone.ShapeGuides)
                {
                    ShapeGuides.Add(g.Clone());
                }
            }

            if (clone.ShapeAdjustHandles!=null)
            {
                ShapeAdjustHandles = new List<ShapeAdjustHandleBase>();
                foreach (var ah in clone.ShapeAdjustHandles)
                {
                    ShapeAdjustHandles.Add(ah.Clone());
                }
            }

            TextBoxRect = clone.TextBoxRect?.Clone();

            ShapePaths = new List<DrawingPath>();
            foreach (var p in clone.ShapePaths)
            {
                ShapePaths.Add(p.Clone());
            }
        }

        public eShapeStyle Style { get; set; }
        /// <summary>
        /// avLst
        /// </summary>
        public List<ShapeGuide> ShapeAdjustValues { get; set; }
        /// <summary>
        /// gdLst  
        /// </summary>
        public List<ShapeGuide> ShapeGuides { get; set; }
        /// <summary>
        /// ahLst 
        /// </summary>
        public List<ShapeAdjustHandleBase> ShapeAdjustHandles { get; set; }
        //cxnLst 
        public List<ShapeConnectionSite> ShapeConnectionSite { get; set; }
        //rect 
        /// <summary>
        /// The rectangle for the text inside the shape.
        /// </summary>
        public TextBoxRect TextBoxRect { get; set; }
        //pathLst
        /// <summary>
        /// Paths to draw the shape
        /// </summary>
        public List<DrawingPath> ShapePaths { get; set; }
        public void Calculate(ExcelShape shape)
        {
            InitCalculatedValues(shape);

            if (ShapeAdjustValues != null)
            {
                var v = shape.GetAdjustmentPointsList(false);
                if (v.Count > 0)
                {
                    var n = shape.GetAdjustmentPointsNames();
                    for (int i = 0; i < n.Count; i++)
                    {
                        _calculatedValues.Add(n[i], v[i]);
                    }
                }
                else
                {
                    foreach (var ap in ShapeAdjustValues)
                    {
                        if (string.IsNullOrEmpty(ap.Formula) == false)
                        {
                            ap.CalculatedValue = CalculateFormula(ap.Formula);
                            _calculatedValues.Add(ap.Name, ap.CalculatedValue);
                        }
                    }
                }
            }
            if (ShapeGuides != null)
            {
                foreach (var g in ShapeGuides)
                {
                    g.CalculatedValue = CalculateFormula(g.Formula);
                    if (_calculatedValues.ContainsKey(g.Name))
                    {
                        _calculatedValues[g.Name] = g.CalculatedValue;
                    }
                    else
                    {
                         _calculatedValues.Add(g.Name, g.CalculatedValue);
                    }
                }
            }
            if(ShapePaths.Count > 0)
            {
                foreach(var item in ShapePaths)
                {
                    shape.GetSizeInPixels(out int width, out int height);
                    var shapeWidth = (double)width * ExcelDrawing.EMU_PER_PIXEL;
                    var shapeHeight = (double)height * ExcelDrawing.EMU_PER_PIXEL;

                    var widthRatio = item.Width.HasValue ? (double)shapeWidth / item.Width : 1D;
                    var heightRatio = item.Height.HasValue ? (double)shapeWidth / item.Height : 1D;
                    item.Width = shapeWidth;
                    item.Height = shapeHeight;

                    foreach (var p in item.Paths)
                    {
                        switch(p.Type)
                        {
                            case PathDrawingType.Close:
                                break;
                            case PathDrawingType.ArcTo:
                                var arc = (ArcTo)p;
                                if (string.IsNullOrEmpty(arc.WidthRadiusName) == false)
                                {
                                    arc.WidthRadius = _calculatedValues[arc.WidthRadiusName];
                                }
                                else
                                {
                                    arc.WidthRadius = (double)((arc.WidthRadius??0D) * widthRatio);
                                }
                                if (string.IsNullOrEmpty(arc.HeightRadiusName) == false)
                                {
                                    arc.HeightRadius = _calculatedValues[arc.HeightRadiusName];
                                }
                                else
                                {
                                    arc.HeightRadius = (double)((arc.HeightRadius ?? 0D) * heightRatio);
                                }
                                if (string.IsNullOrEmpty(arc.StartAngleName) == false) arc.StartAngle = _calculatedValues[arc.StartAngleName];
                                if (string.IsNullOrEmpty(arc.SwingAngleName) == false) arc.SwingAngle = _calculatedValues[arc.SwingAngleName];
                                break;
                            default:
                                var pb = (PathWithCoordinates)p;
                                for(var i=0; i<pb.Coordinates.Count;i++)
                                {
                                    var c = pb.Coordinates[i];
                                    if (string.IsNullOrEmpty(c.XName) == false)
                                    {
                                        c.X = _calculatedValues[c.XName];
                                    }
                                    else
                                    {
                                        c.X = (double)((c.X ?? 0D) * widthRatio);
                                    }

                                    if (string.IsNullOrEmpty(c.YName) == false)
                                    {
                                        c.Y = _calculatedValues[c.YName];
                                    }
                                    else
                                    {
                                        c.Y = (double)((c.Y ?? 0D) * heightRatio);
                                    }
                                }
                                break;
                        }
                    }
                }
            }
        }

        private void InitCalculatedValues(ExcelShape shape)
        {
            shape.GetSizeInPixels(out int width, out int height);
            var shapeWidth = (double)width * ExcelDrawing.EMU_PER_PIXEL;
            var shapeHeight = (double)height * ExcelDrawing.EMU_PER_PIXEL;
            var w = (double)(width * ExcelDrawing.EMU_PER_PIXEL);
            var h = (double)(height * ExcelDrawing.EMU_PER_PIXEL);
            var ls = w < h ? h : w;
            var ss = w > h ? h : w;

            _calculatedValues = new Dictionary<string, double>()
            {
                {"t", 0 },
                {"l", 0 },
                {"w", w },
                {"r", w },
                {"h", h },
                {"b", h },
                {"hc", w/2 },
                {"vc", h/2 },
                {"ls", ls },
                {"ss", ss },
                {"3cd4", 16200000.0d},
                {"3cd8", 8100000.0d},
                {"5cd8", 13500000.0d},
                {"7cd8", 18900000.0d},
                {"cd2", 10800000.0d},
                {"cd4", 5400000.0d},
                {"cd8", 2700000.0d},
                {"hd2", h/2},
                {"hd3", h/3},
                {"hd4", h/4},
                {"hd5", h/5},
                {"hd6", h/6},
                {"hd8", h/8},
                {"wd2", w/2},
                {"wd3", w/3},
                {"wd4", w/4},
                {"wd5", w/5},
                {"wd6", w/6},
                {"wd8", w/8},
                {"wd10", w/10},
                {"wd16", w/16},
                {"wd32", w/32},
                {"ssd2", ss/2 },
                {"ssd4", ss/4 },
                {"ssd6", ss/6 },
                {"ssd8", ss/8 },
                {"ssd16", ss/16 },
                {"ssd32", ss/32 },
            };

        }

        internal double CalculateFormula(string formula)
        {
            var tokens = formula.Split([" "], StringSplitOptions.RemoveEmptyEntries);
            var t = tokens[0];
            switch (t)
            {
                case "val":
                    return GetValue(tokens[1]);
                case "*/":
                    var divBy = GetValue(tokens[3]);
                    if (divBy == 0) return 0;
                    return (GetValue(tokens[1]) * GetValue(tokens[2])) / divBy;
                case "+-":
                    return GetValue(tokens[1]) + GetValue(tokens[2]) - GetValue(tokens[3]);
                case "+/":
                    divBy = GetValue(tokens[3]);
                    if (divBy == 0) return 0;
                    return (GetValue(tokens[1]) + GetValue(tokens[2])) / divBy;
                case "?:":
                    return GetValue(tokens[1]) > 0 ? GetValue(tokens[2]) : GetValue(tokens[3]);
                case "sqrt":
                    return Math.Sqrt(Math.Abs(GetValue(tokens[1])));
                case "abs":
                    return Math.Abs(GetValue(tokens[1]));
                case "min":
                    return Math.Min(GetValue(tokens[1]), GetValue(tokens[2]));
                case "max":
                    return Math.Max(GetValue(tokens[1]), GetValue(tokens[2]));
                case "mod":
                    return Math.Sqrt(Math.Pow((double)GetValue(tokens[1]), 2d) + Math.Pow((double)GetValue(tokens[2]), 2D) + Math.Pow((double)GetValue(tokens[3]), 2D));
                case "pin":
                    //if (y < x), then x = value of this guide else if (y > z), then z
                    double x = GetValue(tokens[1]);
                    double y = GetValue(tokens[2]);
                    double z = GetValue(tokens[3]);

                    return (y < x ? x : y > z ? z : y);
                case "sin":
                    var angleSin = GetValue(tokens[2]) / 60000D;
                    if(angleSin == 0d || angleSin == 180d)
                    {
                        return 0;
                    }
                    var radAngleSin = Math.Sin(Radians(angleSin));
                    return (GetValue(tokens[1]) * radAngleSin);
                case "cos":
                    var angleCos = GetValue(tokens[2]) / 60000D;

                    if (angleCos == 90d || angleCos == 270d)
                    {
                        return 0;
                    }
                    var radAngleCos = Math.Cos(Radians(angleCos));
                    return (GetValue(tokens[1]) * radAngleCos);
                case "tan":
                    var angleTan = GetValue(tokens[2]) / 60000D;

                    if (angleTan == 0d || angleTan == 90d)
                    {
                        //Tan technically undefined
                        return 0;
                    }

                    return (GetValue(tokens[1]) * Math.Tan(Radians(angleTan)));
                case "at2":
                    x = GetValue(tokens[1]);
                    y = GetValue(tokens[2]);
                    double angleRad = Math.Atan2(y, x);     // radians
                    double angleDeg = angleRad * 180.0 / Math.PI; // degrees
                    
                    while (angleDeg < -360)
                        angleDeg += 360;                        // normalize to [0, 360)
                    while (angleDeg > 360)
                    {
                        angleDeg -= 360;
                    }

                    return (angleDeg * 60000D);
                case "cat2":
                    x = GetValue(tokens[1]);
                    y = GetValue(tokens[2]);
                    z = GetValue(tokens[3]);

                    double angleRadCatOrig = Math.Atan2(z, y);     // radians
                    double angleDegCat2Orig = angleRadCatOrig * (180.0 / Math.PI); // degrees

                    if (angleDegCat2Orig == 90d || angleDegCat2Orig == 270d)
                    {
                        return 0;
                    }

                    var dist = (x * Math.Cos(angleRadCatOrig));
                    return dist;
                case "sat2":
                    x = GetValue(tokens[1]);
                    y = GetValue(tokens[2]);
                    z = GetValue(tokens[3]);

                    double angleRadSatOrig = Math.Atan2(z, y);     // radians
                    double angleDegSat2Orig = angleRadSatOrig * (180.0 / Math.PI); // degrees

                    if (angleDegSat2Orig == 0d || angleDegSat2Orig == 180d)
                    {
                        return 0;
                    }

                    var sat2Rad = (x * Math.Sin(angleRadSatOrig));

                    return sat2Rad;
                default:
                    if (_calculatedValues.TryGetValue(t, out var v))
                    {
                        return v;
                    }
                    throw new InvalidOperationException($"Unknown function or variable {{{t}}}");
            }
        }
        private double GetValue(string v)
        {
            if(double.TryParse(v, out var l))
            {
                return l;
            }
            else
            {
                if(_calculatedValues.TryGetValue(v, out var cv))
                {
                    return cv;
                }
                throw new InvalidOperationException($"Unknown variable {{{v}}}");
            }
        }

        internal ShapeDefinition Clone()
        {
            return new ShapeDefinition(this);
        }
        public static double Radians(double angle)
        {
            return (angle / 180) * Math.PI;
        }
    }
}
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Utils;
using System.Xml;

namespace EPPlusImageRenderer.ShapeDefinitions
{
    internal class PresetShapeDefinitions
    {
        static Dictionary<eShapeStyle, ShapeDefinition> _shapeDefinitions = null;
  
        public static Dictionary<eShapeStyle, ShapeDefinition> ShapeDefinitions 
        {
            get
            {
                if(_shapeDefinitions == null)
                {
                    _shapeDefinitions = new Dictionary<eShapeStyle, ShapeDefinition>();
                    Task.Run(() => LoadPresetShapeDefinitionFromXmlAsync()).Wait();
                }
                return _shapeDefinitions;
            }
        }
        public static async Task LoadPresetShapeDefinitionFromXmlAsync()
        {
            var xmlFile = Directory.GetCurrentDirectory() + "\\resource\\presetShapeDefinitions.xml";
            _shapeDefinitions = new Dictionary<eShapeStyle, ShapeDefinition>();
            try
            {
                var ms=new MemoryStream(File.ReadAllBytes(xmlFile));
                var xr = XmlReader.Create(ms, new XmlReaderSettings()
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    IgnoreWhitespace = true,
                    Async = true
                });
                while (await xr.ReadAsync())
                {
                    if (xr.NodeType == XmlNodeType.Element)
                    {
                        if (xr.LocalName != "presetShapeDefinitons")
                        {
                            var item = await LoadPresetShapeDefinitionAsync(xr);
                            _shapeDefinitions.Add(item.Style, item);
                        }
                    }
                }
            }
            catch (Exception ex)
            {   
                ex = ex;
            }
        }

        private static async Task<ShapeDefinition> LoadPresetShapeDefinitionAsync(XmlReader xr)
        {
            if (Enum.TryParse<eShapeStyle>(xr.LocalName, true, out var style))
            {
                var psd = new ShapeDefinition()
                {
                    Style = style
                };
                await LoadFromXmlAsync(psd, xr);
                return psd;
            }
            throw new InvalidOperationException();
        }
        private static async Task LoadFromXmlAsync(ShapeDefinition psd, XmlReader xr)
        {
            while (await xr.ReadAsync())
            {
                if (xr.NodeType == XmlNodeType.Element)
                {
                    switch (xr.LocalName)
                    {
                        case "avLst":
                            psd.ShapeAdjustValues = await LoadShapeGuidesAsync(xr);
                            break;
                        case "gdLst":
                            psd.ShapeGuides = await LoadShapeGuidesAsync(xr);
                            break;
                        case "ahLst":
                            psd.ShapeAdjustHandles = await LoadAdjustHandleAsync(xr);
                            break;
                        case "cxnLst":
                            psd.ShapeConnectionSite = await LoadConnectionLstAsync(xr);
                            break;
                        case "rect":
                            psd.TextBoxRect = new TextBoxRect() { Top = xr.GetAttribute("t"), Bottom = xr.GetAttribute("b"), Left = xr.GetAttribute("l"), Right = xr.GetAttribute("r") };
                            break;
                        case "pathLst":
                            psd.ShapePaths = LoadShapePaths(xr);
                            break;
                    }
                }
                else if (xr.NodeType == XmlNodeType.EndElement && xr.LocalName.Equals(psd.Style.ToString(), StringComparison.InvariantCultureIgnoreCase))
                {
                    return;
                }
            }
        }

        private static List<DrawingPath> LoadShapePaths(XmlReader xr)
        {
            var list = new List<DrawingPath>();
            while (xr.Read())
            {
                if (xr.LocalName == "path" && xr.NodeType == XmlNodeType.Element)
                {
                    list.Add(new DrawingPath(xr));
                }
                else if (xr.IsEndElementWithName("pathLst"))
                {
                    break;
                }
            }
            return list;
        }

        private static async Task<List<ShapeConnectionSite>> LoadConnectionLstAsync(XmlReader xr)
        {
            List<ShapeConnectionSite> shapeConnectionSite = new();

            while (await xr.ReadAsync() && (xr.NodeType != XmlNodeType.EndElement && xr.LocalName != "cxnLst"))
            {
                var newConnection = new ShapeConnectionSite();

                var attrStr = xr.GetAttribute("ang");

                newConnection.Angle = xr.GetAttribute("ang");
                await xr.ReadAsync();

                newConnection.PositionCoordinate = new ShapePositionCoordinate() { X = xr.GetAttribute("x"), Y = xr.GetAttribute("y") };

                shapeConnectionSite.Add(newConnection);
                await xr.ReadAsync();

                if (xr.LocalName == "cxnLst" && xr.NodeType == XmlNodeType.EndElement)
                {
                    break;
                }
            }

            return shapeConnectionSite;
        }

        //private static int? GetNumericValueFromVariable(string? variableName)
        //{
        //    var wasParsed = int.TryParse(variableName, CultureInfo.InvariantCulture, out int res);

        //    if (wasParsed)
        //    {
        //        return res;
        //    }
        //    else
        //    {
        //        //TODO: Define variable lookup
        //        //variableDict["variableName"]
        //        return null;
        //    }
        //}

        private static async Task<List<ShapeAdjustHandleBase>> LoadAdjustHandleAsync(XmlReader xr)
        {
            var l = new List<ShapeAdjustHandleBase>();
            var name = xr.LocalName;
            while (await xr.ReadAsync())
            {
                if (xr.NodeType == XmlNodeType.Element)
                {
                    switch (xr.LocalName)
                    {
                        case "ahXY":
                            l.Add(ShapeAdjustHandleBase.CreateXy(xr));
                            break;
                        case "ahPolar":
                            l.Add(ShapeAdjustHandleBase.CreatePolar(xr));
                            break;
                        case "pos":
                            l[l.Count - 1].PositionCoordinate = new ShapePositionCoordinate() { X = xr.GetAttribute("x"), Y = xr.GetAttribute("y") };
                            break;
                    }
                }
                else if (xr.NodeType == XmlNodeType.EndElement && xr.LocalName == name)
                {
                    break;
                }
            }
            return l;
        }

        private static async Task<List<ShapeGuide>> LoadShapeGuidesAsync(XmlReader xr)
        {
            var l = new List<ShapeGuide>();
            var name = xr.LocalName;
            while (await xr.ReadAsync())
            {
                if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "gd")
                {
                    l.Add(new ShapeGuide() { Name = xr.GetAttribute("name"), Formula = xr.GetAttribute("fmla") });
                }

                if (xr.NodeType == XmlNodeType.EndElement && xr.LocalName == name)
                {
                    break;
                }
            }
            return l;
        }
    }
}
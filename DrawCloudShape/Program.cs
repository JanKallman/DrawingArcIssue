using EPPlusImageRenderer.ShapeDefinitions;
using OfficeOpenXml.Drawing;
using System.Diagnostics;
using System.Globalization;
using System.Text;

var ci = Thread.CurrentThread.CurrentCulture;
Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
await PresetShapeDefinitions.LoadPresetShapeDefinitionFromXmlAsync();

//This sample is hard coded aginst the Cloud shape in the presetShapeDefinitions.xml.
//As the height and width in the definintion is 43200 we just divide the values with 100 for a svg size of 432 pixels.
var cloudShape = PresetShapeDefinitions.ShapeDefinitions[OfficeOpenXml.Drawing.eShapeStyle.Cloud].Clone();
StringBuilder svgSb = new StringBuilder();
svgSb.AppendLine("<svg width=\"432\" height=\"432\" xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xml:space=\"preserve\" overflow=\"visible\" viewbox=\"0,0,500,500\">");
foreach(var path in cloudShape.ShapePaths)
{
    var width = path.Width; //43200 for a Cloud shape
    var height = path.Height;
    svgSb.Append(" <path d=\"");
    PathDrawingType? pCmd = null;
    double cx = 0, cy = 0;
    foreach(var command in path.Paths)
    {
        switch(command.Type)
        {
            case PathDrawingType.MoveTo:
                var moveTo=command as MoveTo;
                svgSb.Append("M");
                foreach (var c in moveTo.Coordinates)
                {
                    svgSb.Append($"{c.X/100},{c.Y/100}"); //Divide with 100 for size 432
                    cx = (c.X??0)/100; //Divide with 100 for size 432
                    cy = (c.Y??0)/100; //Divide with 100 for size 432
                }
                break;
            case PathDrawingType.ArcTo:
                var arcTo=command as ArcTo;
                var wR = arcTo.WidthRadius.Value / 100;             //Divide with 100 for size 432
                var hR = arcTo.HeightRadius.Value / 100;            //Divide with 100 for size 432
                var startAngle = arcTo.StartAngle.Value / 60000D;   //Angles are in 60000th of a degree.
                var swingAngle = arcTo.SwingAngle.Value / 60000D;   //Angles are in 60000th of a degree.

                var startAngleRad = ShapeDefinition.Radians(startAngle);
                var endAngleRad = ShapeDefinition.Radians(startAngle + swingAngle);

                //************************************************************************************
                //Here is the calculation that needs to be corrected.
                //It is not clear how to calculate the end coordinates for the ArcTo element.
                //As the end coordinates are used as input to the next arc the two paths of the cloud does not align.
                //************************************************************************************
                double centerX, centerY;
                centerX = cx - wR * Math.Cos(startAngleRad);
                centerY = cy - hR * Math.Sin(startAngleRad);

                var endX = centerX + wR * Math.Cos(endAngleRad);
                var endY = centerY + hR * Math.Sin(endAngleRad);

                //************************************************************************************

                svgSb.Append($"A{wR} {hR} 0 0 {(swingAngle < 0 ? 0 : 1)} {endX} {endY}");
                cx = endX;
                cy = endY;
                break;
            case PathDrawingType.Close:
                svgSb.Append($"Z");
                break;
        }
        pCmd = command.Type;
    }
    svgSb.Append("\" fill=\"transparent\" stroke=\"#0a3041\" stroke-width=\"1.125\" stroke-miterlimit=\"8\" />"); 
}

svgSb.AppendLine("</svg>");
File.WriteAllText("cloud.svg", svgSb.ToString());
var svgFile = new FileInfo("cloud.svg");
Console.WriteLine($"Svg written to: {svgFile.FullName}");
Console.ReadKey();
//Process.Start("notepad.exe", svgFile.FullName); //Uncomment to open in notepad.
Thread.CurrentThread.CurrentCulture = ci;

using System.Diagnostics;

namespace EPPlusImageRenderer.ShapeDefinitions
{
    /// <summary>
    /// 20.1.9.11 Ecma part 1
    /// 20.1.10.76 - Preset Text Shape Types
    /// </summary>
    [DebuggerDisplay("{Name}-{Formula}={CalculatedValue}")]
    public class ShapeGuide
    {
        public string Name { get; set; }
        public string Formula { get; set; }
        public double CalculatedValue 
        { 
            get; 
            set; 
        }

        internal ShapeGuide Clone()
        {
            return new ShapeGuide() { Name=Name, Formula=Formula, CalculatedValue=CalculatedValue};
        }
    }
}

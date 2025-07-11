
namespace EPPlusImageRenderer.ShapeDefinitions
{
    public class TextBoxRect
    {
        public object Left { get; set; }
        public object Right { get; set; }
        public object Top { get; set; }
        public object Bottom { get; set; }

        internal TextBoxRect Clone()
        {
            return new TextBoxRect()
            {
                Left = Left,
                Right = Right,
                Top = Top,
                Bottom = Bottom
            };   
        }
    }
}
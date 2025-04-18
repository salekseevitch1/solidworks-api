namespace CADBooster.SolidDna
{
    public class Outline2D : Outline
    {
        public Outline2D(XYZ min, XYZ max) : base(min, max)
        {

        }

        public double Width => Max.X - Min.X;
        public double Height => Max.Y - Min.Y;

        /// <summary>
        /// Converts the outline to a vector
        /// </summary>
        /// <returns></returns>
        public XYZ AsVector() => new XYZ(Width, Height, 0);
    }
}
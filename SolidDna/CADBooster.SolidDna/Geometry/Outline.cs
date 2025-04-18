namespace CADBooster.SolidDna
{
    public class Outline
    {
        public XYZ Min { get; set; }
        public XYZ Max { get; set; }

        public XYZ Center => (Min + Max) / 2;

        public Outline(XYZ min, XYZ max)
        {
            Min = min;
            Max = max;
        }
    }
}
namespace CADBooster.SolidDna
{
    public class MarginSettings
    {
        public double Left { get; set; }
        public double Right { get; set; }
        public double Top { get; set; }
        public double Bottom { get; set; }

        public static MarginSettings Zero
            => new MarginSettings(0, 0, 0, 0);

        public MarginSettings(double left, double top, double right, double bottom)
        {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }
    }
}

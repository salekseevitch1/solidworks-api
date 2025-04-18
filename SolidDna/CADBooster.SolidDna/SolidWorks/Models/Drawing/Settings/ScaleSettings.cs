namespace CADBooster.SolidDna
{
    public class ScaleSettings
    {
        public int Numerator { get; set; }
        public int Denominator { get; set; }

        public static ScaleSettings FullScale 
            => new ScaleSettings(1, 1);

        public ScaleSettings(int numerator, int denominator)
        {
            Numerator = numerator;
            Denominator = denominator;
        }
    }
}
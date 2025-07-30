using CADBooster.SolidDna;

public class CircleParameters
{
    public XYZ Center { get; set; } = XYZ.Zero;
    public XYZ Axis { get; set; } = XYZ.Zero;
    public double Radius { get; set; } = double.NaN;
}
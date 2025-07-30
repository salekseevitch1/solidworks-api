using CADBooster.SolidDna;
using SolidWorks.Interop.sldworks;
using System;

public static class CurveExtensions
{
    public static CircleParameters GetCircleParameters(this Curve curve)
    {
        if (curve.IsCircle() == false)
            throw new InvalidOperationException("Curve is not a circle.");

        var parameters = (double[])curve.CircleParams;

        return new CircleParameters
        {
            Center = new XYZ(parameters[0], parameters[1], parameters[2]),
            Axis = new XYZ(parameters[3], parameters[4], parameters[5]),
            Radius = parameters[6],
        };
    }
}
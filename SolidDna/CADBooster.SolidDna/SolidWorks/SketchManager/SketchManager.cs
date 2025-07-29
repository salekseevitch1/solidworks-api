using SolidWorks.Interop.sldworks;

namespace CADBooster.SolidDna.SketchManager;

public class SketchManager(global::SolidWorks.Interop.sldworks.SketchManager comObject)
    : SolidDnaObject<global::SolidWorks.Interop.sldworks.SketchManager>(comObject)
{
    public SketchPoint CreatePoint(XYZ point)
    {
        UnsafeObject.InsertSketch(true);

        UnsafeObject.AddToDB = true;

        var sketchPoint = UnsafeObject.CreatePoint(
            point.X,
            point.Y,
            point.Z);

        UnsafeObject.AddToDB = false;

        UnsafeObject.InsertSketch(false);

        return sketchPoint;
    }

    public SketchSegment CreateCircle(XYZ center, double diameter)
    {
        UnsafeObject.InsertSketch(true);

        UnsafeObject.AddToDB = true;

        var pointOnCircle = center.MoveAlongVector(XYZ.BasisX, diameter / 2);

        var sketchSegment = UnsafeObject.CreateCircle(
            center.X,
            center.Y,
            center.Z,
            pointOnCircle.X,
            pointOnCircle.Y,
            pointOnCircle.Z);

        UnsafeObject.AddToDB = false;

        UnsafeObject.InsertSketch(false);

        return sketchSegment;
    }
}
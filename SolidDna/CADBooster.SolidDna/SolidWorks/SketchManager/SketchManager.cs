using SolidWorks.Interop.sldworks;

namespace CADBooster.SolidDna.SketchManager;

public class SketchManager(global::SolidWorks.Interop.sldworks.SketchManager comObject)
    : SolidDnaObject<global::SolidWorks.Interop.sldworks.SketchManager>(comObject)
{
    public SketchPoint CreatePoint(double[] point)
    {
        UnsafeObject.InsertSketch(true);

        UnsafeObject.AddToDB = true;

        var sketchPoint = UnsafeObject.CreatePoint(
            point[0],
            point[1],
            point[2]);

        UnsafeObject.AddToDB = false;

        UnsafeObject.InsertSketch(false);

        return sketchPoint;
    }

    public SketchSegment CreateCircle(double[] center, double diameter)
    {
        UnsafeObject.InsertSketch(true);

        UnsafeObject.AddToDB = true;

        var pointOnCircle = new[]
        {
            center[0] + (diameter / 2),
            center[1],
            center[2]
        };

        var sketchSegment = UnsafeObject.CreateCircle(
            center[0],
            center[1],
            center[2],
            pointOnCircle[0],
            pointOnCircle[1],
            pointOnCircle[2]);

        UnsafeObject.AddToDB = false;

        UnsafeObject.InsertSketch(false);

        return sketchSegment;
    }
}
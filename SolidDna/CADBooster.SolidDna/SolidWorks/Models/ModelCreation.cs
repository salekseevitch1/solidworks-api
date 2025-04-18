namespace CADBooster.SolidDna
{
    public class ModelCreation
    {
        private readonly Model _model;

        public ModelCreation(Model model)
        {
            _model = model;
        }

        public void CreateLine(XYZ start, XYZ end)
        {
            _model.UnsafeObject.CreateLine2(
                start.X,
                start.Y,
                start.Z,

                end.X,
                end.Y,
                end.Z);
        }
    }
}
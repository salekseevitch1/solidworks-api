using SolidWorks.Interop.swconst;

namespace CADBooster.SolidDna
{
    public class DrawingDocumentCreation
    {
        private readonly DrawingDocument _drawingDocument;

        public DrawingDocumentCreation(DrawingDocument drawingDocument)
        {
            _drawingDocument = drawingDocument;
        }

        public DrawingView CreateSection(
            XYZ position,
            string sectionName,
            swCreateSectionViewAtOptions_e options,
            double depth = 0)
        {
            var drawing = _drawingDocument.UnsafeObject;

            var view = drawing.CreateSectionViewAt5(
                position.X,
                position.Y,
                position.Z,
                sectionName,
                (int)options,
                (object)null,
                0);

            return new DrawingView(view);
        }
    }
}
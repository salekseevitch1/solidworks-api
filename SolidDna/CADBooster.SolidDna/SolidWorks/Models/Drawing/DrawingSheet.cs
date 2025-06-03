using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace CADBooster.SolidDna
{
    /// <summary>
    /// A sheet of a drawing
    /// </summary>
    public class DrawingSheet : SolidDnaObject<Sheet>
    {
        #region Private Members

        /// <summary>
        /// The parent drawing document of this sheet
        /// </summary>
        private readonly DrawingDocument mDrawingDoc;

        #endregion

        #region Public Properties

        /// <summary>
        /// The sheet name
        /// </summary>
        public string SheetName => BaseObject.GetName();

        #endregion

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="comObject">The underlying COM object</param>
        /// <param name="drawing">The parent drawing document</param>
        public DrawingSheet(Sheet comObject, DrawingDocument drawing) : base(comObject)
        {
            mDrawingDoc = drawing;
        }

        #endregion

        /// <summary>
        /// Activates this sheet
        /// </summary>
        public bool Activate() => mDrawingDoc.ActivateSheet(SheetName);

        public Size GetSize(out swDwgPaperSizes_e paperSize)
        {
            var width = 0d;
            var height = 0d;

            paperSize = (swDwgPaperSizes_e)BaseObject.GetSize(ref width, ref height);

            return new Size(width, height);
        }

        public XYZ GetCenter()
        {
            var viewsSheetSize = GetSize(out _);

            var centerX = viewsSheetSize.Width / 2;
            var centerY = viewsSheetSize.Height / 2;

            return new XYZ(centerX, centerY, 0);
        }

        public void GetScale(out int numerator, out int denominator)
        {
            var properties = (double[])UnsafeObject.GetProperties2();

            numerator = (int)properties[2];
            denominator = (int)properties[3];
        }
    }
}

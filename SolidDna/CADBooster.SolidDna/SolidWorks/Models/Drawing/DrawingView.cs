using SolidWorks.Interop.sldworks;

namespace CADBooster.SolidDna
{
    /// <summary>
    /// A view of a drawing
    /// </summary>
    public class DrawingView : SolidDnaObject<View>
    {
        #region Public Properties

        /// <summary>
        /// The drawing view type
        /// </summary>
        public DrawingViewType ViewType => (DrawingViewType)BaseObject.Type;

        /// <summary>
        /// The name of the view
        /// </summary>
        public string Name => BaseObject.Name;

        /// <summary>
        /// The X position of the view origin with respect to the drawing sheet origin
        /// </summary>
        public double PositionX => ((double[])BaseObject.Position)[0];

        /// <summary>
        /// The Y position of the view origin with respect to the drawing sheet origin
        /// </summary>
        public double PositionY => ((double[])BaseObject.Position)[1];

        public XYZ Position => new XYZ(PositionX, PositionY, 0);

        public Model ReferencedModel => new Model(UnsafeObject.ReferencedDocument);

        /// <summary>
        /// The bounding box of the view
        /// </summary>
        public BoundingBox BoundingBox
        {
            get
            {
                var box = (double[])BaseObject.GetOutline();
                return new BoundingBox(box[0], box[1], box[2], box[3]);
            }
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="comObject">The underlying COM object</param>
        public DrawingView(View comObject) : base(comObject)
        {
        }

        #endregion

        #region Public Members

        public XYZ GetNormal()
        {
            var transform = UnsafeObject.ModelToViewTransform;

            var transformArray = (double[])transform.ArrayData;

            return new XYZ(transformArray[6], transformArray[7], transformArray[8]);
        }

        public Outline2D GetOutline()
        {
            var box = (double[])BaseObject.GetOutline();

            var min = new XYZ(box[0], box[1], 0);
            var max = new XYZ(box[2], box[3], 0);


            return new Outline2D(min, max);
        }

        #endregion
    }
}

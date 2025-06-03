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
        
        /// <summary>
        /// Gets the position
        /// </summary>
        public XYZ Position => new XYZ(PositionX, PositionY, 0);

        /// <summary>
        /// Gets the model referenced by the associated document.
        /// </summary>
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

        /// <summary>
        /// Retrieves the normal vector of the associated object in the view coordinate system.
        /// </summary>
        /// <remarks>The normal vector is derived from the transformation matrix of the object.</remarks>
        /// <returns>An <see cref="XYZ"/> object representing the normal vector in the view coordinate system.</returns>
        public XYZ GetNormal()
        {
            var transform = UnsafeObject.ModelToViewTransform;

            var transformArray = (double[])transform.ArrayData;

            return new XYZ(transformArray[6], transformArray[7], transformArray[8]);
        }

        /// <summary>
        /// Retrieves the 2D outline of the object as an <see cref="Outline2D"/> instance.
        /// </summary>
        /// <remarks>The outline is defined by a rectangular bounding box in 2D space, represented by its
        /// minimum and maximum coordinates. The Z-coordinate of the resulting points is set to 0.</remarks>
        /// <returns>An <see cref="Outline2D"/> object representing the 2D bounding box of the object.</returns>
        public Outline2D GetOutline()
        {
            var box = (double[])BaseObject.GetOutline();

            var min = new XYZ(box[0], box[1], 0);
            var max = new XYZ(box[2], box[3], 0);

            return new Outline2D(min, max);
        }

        /// <summary>
        /// Sets the scale value for the current object.
        /// </summary>
        /// <remarks>This method updates the scale of the underlying object. Ensure that the provided
        /// <paramref name="scale"/>  is within the acceptable range for the object's scale property to avoid unexpected
        /// behavior.</remarks>
        /// <param name="scale">The scale value to set. Must be a valid double precision number.</param>
        public void SetScale(double scale) 
            => UnsafeObject.ScaleDecimal = scale;


        /// <summary>
        /// Sets the position of the object using the specified <see cref="XYZ"/> instance.
        /// </summary>
        /// <remarks>This method updates the object's position based on the array data contained in the
        /// provided <see cref="XYZ"/> instance. Ensure that the <paramref name="position"/> is properly initialized
        /// before calling this method.</remarks>
        /// <param name="position">The <see cref="XYZ"/> instance representing the new position.  The <paramref name="position"/> must contain
        /// valid array data.</param>
        public void SetPosition(XYZ position)
            => UnsafeObject.Position = position.ArrayData;

        #endregion
    }
}

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;

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
        public double[] Position => [PositionX, PositionY, 0];

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
        public void SetPosition(double[] position)
            => UnsafeObject.Position = position;

        #endregion

        public List<T> GetEntitiesByCondition<T>(Func<T, bool> condition)
        {
            var components = (object[])UnsafeObject.GetVisibleComponents();

            var tElements = new List<T>();

            foreach (Component2 component2 in components)
            {
                var entities = (object[])UnsafeObject.GetVisibleEntities2(
                    component2,
                    (int)swViewEntityType_e.swViewEntityType_Edge);

                var edges = new List<IEdge>();

                foreach (IEntity entity in entities)
                {
                    if (entity is not T tEntity)
                        continue;

                    if (condition(tEntity))
                        tElements.Add(tEntity);
                }
            }

            return tElements;
        }
    }
}

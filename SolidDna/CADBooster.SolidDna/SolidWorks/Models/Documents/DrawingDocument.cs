using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CADBooster.SolidDna
{
    /// <summary>
    /// Exposes all Drawing Document calls from a <see cref="Model"/>
    /// </summary>
    public class DrawingDocument
    {
        #region Protected Members

        /// <summary>
        /// The base model document. Note we do not dispose of this (the parent Model will)
        /// </summary>
        protected DrawingDoc mBaseObject;

        #endregion

        #region Public Properties

        /// <summary>
        /// The raw underlying COM object
        /// WARNING: Use with caution. You must handle all disposal from this point on
        /// </summary>
        public DrawingDoc UnsafeObject => mBaseObject;

        public string Title => ((IModelDoc2)UnsafeObject).GetTitle();

        public DrawingDocumentCreation Creation => new DrawingDocumentCreation(this);

        #endregion

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public DrawingDocument(DrawingDoc model)
        {
            mBaseObject = model;
        }

        #endregion

        #region Feature Methods

        /// <summary>
        /// Get the <see cref="ModelFeature"/> of the item in the feature tree based on its name. Returns the actual model feature.
        /// </summary>
        /// <param name="featureName">Name of the feature</param>
        /// <returns>The <see cref="ModelFeature"/> for the named feature</returns>
        public ModelFeature GetFeatureByName(string featureName)
        {
            // Wrap any error
            return SolidDnaErrors.Wrap(() => GetModelFeatureByNameOrNull(featureName),
                SolidDnaErrorTypeCode.SolidWorksModel,
                SolidDnaErrorCode.SolidWorksModelAssemblyGetFeatureByNameError);
        }

        /// <summary>
        /// Get the <see cref="ModelFeature"/> of the item in the feature tree based on its name and perform a function on it.
        /// </summary>
        /// <param name="featureName">Name of the feature</param>
        /// <param name="function">The function to perform on this feature</param>
        /// <returns>The <see cref="ModelFeature"/> for the named feature</returns>
        public T GetFeatureByName<T>(string featureName, Func<ModelFeature, T> function)
        {
            // Wrap any error
            return SolidDnaErrors.Wrap(() =>
            {
                // Create feature
                using (var modelFeature = GetModelFeatureByNameOrNull(featureName))
                {
                    // Run function
                    return (T)function.Invoke(modelFeature);
                }
            },
                SolidDnaErrorTypeCode.SolidWorksModel,
                SolidDnaErrorCode.SolidWorksModelAssemblyGetFeatureByNameError);
        }

        /// <summary>
        /// Get the <see cref="ModelFeature"/> of the item in the feature tree based on its name and perform an action on it.
        /// </summary>
        /// <param name="featureName">Name of the feature</param>
        /// <param name="action">The action to perform on this feature</param>
        /// <returns>The <see cref="ModelFeature"/> for the named feature</returns>
        public void GetFeatureByName(string featureName, Action<ModelFeature> action)
        {
            // Wrap any error
            SolidDnaErrors.Wrap(() =>
            {
                // Create feature
                using (var modelFeature = GetModelFeatureByNameOrNull(featureName))
                {
                    // Run action
                    action(modelFeature);
                }
            },
                SolidDnaErrorTypeCode.SolidWorksModel,
                SolidDnaErrorCode.SolidWorksModelAssemblyGetFeatureByNameError);
        }

        /// <summary>
        /// Get the <see cref="ModelFeature"/> of the item in the feature tree based on its name.
        /// Returns the actual model feature or null when not found.
        /// </summary>
        /// <param name="featureName"></param>
        /// <returns></returns>
        private ModelFeature GetModelFeatureByNameOrNull(string featureName)
        {
            var feature = (Feature)mBaseObject.FeatureByName(featureName);
            return feature == null ? null : new ModelFeature(feature);
        }

        #endregion

        #region Sheet Methods

        /// <summary>
        /// Activates the specified drawing sheet
        /// </summary>
        /// <param name="sheetName">Name of the sheet</param>
        /// <returns>True if the sheet was activated, false if SOLIDWORKS generated an error</returns>
        public bool ActivateSheet(string sheetName) => mBaseObject.ActivateSheet(sheetName);

        /// <summary>
        /// Gets the name of the currently active sheet
        /// </summary>
        /// <returns></returns>
        public string CurrentActiveSheet()
        {
            using (var sheet = new DrawingSheet((Sheet)mBaseObject.GetCurrentSheet(), this))
            {
                return sheet.SheetName;
            }
        }

        /// <summary>
        /// Retrieves the currently active drawing sheet.
        /// </summary>
        /// <remarks>The returned <see cref="DrawingSheet"/> object is initialized based on the current
        /// sheet in the underlying base object. Ensure that the base object is properly configured before calling this
        /// method.</remarks>
        /// <returns>An instance of <see cref="DrawingSheet"/> representing the active sheet. If no sheet is active, the method
        /// may return a default or placeholder sheet.</returns>
        public DrawingSheet GetActiveSheet()
        {
            return new DrawingSheet((Sheet)mBaseObject.GetCurrentSheet(), this);
        }

        /// <summary>
        /// Gets the sheet names of the drawing
        /// </summary>
        /// <returns></returns>
        public string[] SheetNames() => (string[])mBaseObject.GetSheetNames();

        public void ForEachSheet(Action<DrawingSheet> sheetsCallback)
        {
            // Get each sheet name
            var sheetNames = SheetNames();

            // Get all sheet names
            foreach (var sheetName in sheetNames)
            {
                // Get instance of sheet
                using (var sheet = new DrawingSheet(mBaseObject.Sheet[sheetName], this))
                {
                    // Callback
                    sheetsCallback(sheet);
                }
            }
        }

        public DrawingSheet NewSheet(
            string name,
            swDwgPaperSizes_e paperSize,
            swDwgTemplates_e template,
            ScaleSettings scale = null,
            MarginSettings margin = null)
        {
            scale = scale ?? ScaleSettings.FullScale;
            margin = margin ?? MarginSettings.Zero;

            mBaseObject.NewSheet4(
                // Name to be given to the new drawing sheet
                Name: name,
                // Size of paper as defined in swDwgPaperSizes_e;
                // valid only if TemplateIn is swDwgTemplates_e.swDwgTemplateNone
                PaperSize: (int)paperSize,
                // Template as defined in swDwgTemplates_e
                TemplateIn: (int)template,
                // Scale numerator
                Scale1: scale.Numerator,
                // Scale denominator
                Scale2: scale.Denominator,
                // True for first angle projection, false for third angle projection
                FirstAngle: false,
                // Name of custom template with full directory path;
                // valid only if TemplateIn is set
                // to swDwgTemplates_e.swDwgTemplateCustom
                TemplateName: string.Empty,
                // Paper width; valid only if TemplateIn is set
                // to swDwgTemplates_e.swDwgTemplateNone
                // or PaperSize is set to swDwgPaperSizes_e.swDwgPapersUserDefined
                Width: 0,
                // Paper height; valid only if TemplateIn is set
                // to swDwgTemplates_e.swDwgTemplateNone or PaperSize
                // is set to swDwgPaperSizes_e.swDwgPapersUserDefined
                Height: 0,
                // Name of view containing the model from which
                // to get custom property values
                PropertyViewName: "",
                // Zone area left margin; distance from drawing sheet's left edge
                ZoneLeftMargin: margin.Left,
                // Zone area right margin; distance from drawing sheet's right edge
                ZoneRightMargin: margin.Right,
                // Zone area top margin; distance from drawing sheet's top edge
                ZoneTopMargin: margin.Top,
                // Zone area bottom margin; distance from drawing sheet's bottom edge
                ZoneBottomMargin: margin.Bottom,
                // Number of zone rows in the zone area of this sheet (see Remarks)
                ZoneRow: 0,
                // Number of zone columns in the zone area of this sheet (see Remarks)
                ZoneCol: 0
            );

            // Remarks
            // The drawing sheet can be created with zones that annotations
            // in other views can reference. Each zone is referenced by
            // an alphanumeric label that is defined using the Zone Editor.
            // See the SOLIDWORKS Help for more information about drawing sheet zones.
            // 
            // (ZoneRow x ZoneCol) is the total number of zones in the zone
            // area of this drawing sheet. The zone area is specified
            // by ZoneLeftMargin, ZoneRightMargin, ZoneTopMargin, and ZoneBottomMargin.

            return new DrawingSheet(mBaseObject.Sheet[name], this);
        }

        #endregion

        #region View Methods

        /// <summary>
        /// Activates the specified drawing view
        /// </summary>
        /// <param name="viewName">Name of the drawing view</param>
        /// <returns>True if successful, false if not</returns>
        public bool ActivateView(string viewName) => mBaseObject.ActivateView(viewName);

        /// <summary>
        /// Rotates the view so the selected line in the view is horizontal
        /// </summary>
        public void AlignViewHorizontally() => mBaseObject.AlignHorz();

        /// <summary>
        /// Rotates the view so the selected line in the view is vertical
        /// </summary>
        public void AlignViewVertically() => mBaseObject.AlignVert();

        /// <summary>
        /// Gets all the views of the drawing
        /// </summary>
        /// <param name="viewsCallback">The callback containing all views</param>
        public void Views(Action<List<DrawingView>> viewsCallback)
        {
            // List of all views
            var views = new List<DrawingView>();

            // Get all views as an array of arrays
            var sheetArray = (object[])mBaseObject.GetViews();

            // Get all views
            foreach (object[] viewArray in sheetArray)
                foreach (View view in viewArray)
                    views.Add(new DrawingView((View)view));

            try
            {
                // Callback
                viewsCallback(views);
            }
            finally
            {
                // Dispose all views
                views.ForEach(view => view.Dispose());
            }
        }

        public DrawingView CreateDrawViewFromModelView(
            PartDocument part,
            string viewName) 
            => CreateDrawViewFromModelView(part, viewName, XYZ.Zero);

        public DrawingView CreateDrawViewFromModelView(
            PartDocument part,
            string viewName,
            XYZ point)
        {
            var partPath = ((IModelDoc2)part.UnsafeObject).GetPathName();

            var view = UnsafeObject.CreateDrawViewFromModelView3(
                ModelName: partPath,
                ViewName: viewName,
                LocX: point.X,
                LocY: point.Y,
                LocZ: point.Z);

            return new DrawingView(view);
        }

        public DrawingView CreateDrawViewFromModelView(
            string partPath,
            ViewType viewType,
            XYZ point,
            string viewName = "")
        {
            var view = UnsafeObject.CreateDrawViewFromModelView3(
                ModelName: partPath,
                ViewName: $"*{viewType.ToString()}",
                LocX: point.X,
                LocY: point.Y,
                LocZ: point.Z);

            if (!string.IsNullOrEmpty(viewName))
                view.SetName2(viewName);

            return new DrawingView(view);
        }

        public DrawingView GetView(string sheetName, string viewName)
        {
            var views = UnsafeObject.Sheet[sheetName].GetViews() as object[];

            var view = views?
                .Cast<View>()
                .FirstOrDefault(it => string.Equals(it.Name, viewName));

            return new DrawingView(view);
        }

        #endregion

        #region Balloon Methods

        /// <summary>
        /// Automatically inserts BOM balloons in selected drawing views
        /// </summary>
        /// <param name="options">The balloon options</param>
        /// <param name="onSuccess">Callback containing all created notes if successful</param>
        /// <returns>An array of the created <see cref="Note"/> objects</returns>
        /// <remarks>
        ///     This method automatically inserts BOM balloons for bill of materials 
        ///     items in selected drawing views. If a drawing sheet is selected, BOM 
        ///     balloons are automatically inserted for all the drawing views on that 
        ///     drawing sheet. 
        ///     
        ///     To automatically insert BOM balloons, select one or more views or sheets
        ///     for which to automatically create BOM balloons, then call this method.
        /// </remarks>
        public void AutoBalloon(AutoBalloonOptions options, Action<Note[]> onSuccess = null)
        {
            // Create native options
            var nativeOptions = mBaseObject.CreateAutoBalloonOptions();

            // Fill options
            nativeOptions.CustomSize = options.CustomSize;
            nativeOptions.EditBalloonOption = (int)options.EditBalloonOptions;
            nativeOptions.EditBalloons = options.EditBalloons;
            nativeOptions.FirstItem = options.FirstItem;
            nativeOptions.IgnoreMultiple = options.IgnoreMultiple;
            nativeOptions.InsertMagneticLine = options.InsertMagneticLine;
            nativeOptions.ItemNumberIncrement = options.ItemNumberIncrement;
            nativeOptions.ItemNumberStart = options.ItemNumberStart;
            nativeOptions.ItemOrder = (int)options.ItemOrder;
            nativeOptions.Layername = options.LayerName;
            nativeOptions.Layout = (int)options.Layout;
            nativeOptions.LeaderAttachmentToFaces = options.LeaderAttachmentToFaces;
            nativeOptions.LowerText = options.LowerText;
            nativeOptions.LowerTextContent = (int)options.LowerTextContent;
            nativeOptions.ReverseDirection = options.ReverseDirection;
            nativeOptions.Size = (int)options.Size;
            nativeOptions.Style = (int)options.Style;
            nativeOptions.UpperText = options.UpperText;
            nativeOptions.UpperTextContent = (int)options.UpperTextContent;

            // Create the notes
            var nativeNotes = (object[])mBaseObject.AutoBalloon5(nativeOptions);

            // If we have a callback, and have notes...
            if (onSuccess != null && nativeNotes?.Length > 0)
            {
                // Create all note classes
                var notes = nativeNotes.Select(f => new Note((INote)f)).ToArray();

                // Inform listeners
                onSuccess.Invoke(notes);

                // Dispose them
                foreach (var note in notes)
                    note.Dispose();
            }
        }

        #endregion

        #region Dimension Methods

        /// <summary>
        /// Adds a chamfer dimension to the selected edges
        /// </summary>
        /// <param name="x">X dimension</param>
        /// <param name="y">Y dimension</param>
        /// <param name="z">Z dimension</param>
        /// <returns>The chamfer <see cref="ModelDisplayDimension"/> if successful. Null if not.</returns>
        /// <remarks>Make sure to select the 2 edges of the chamfer before running this command</remarks>
        public ModelDisplayDimension AddChamferDimension(double x, double y, double z)
            => new ModelDisplayDimension((IDisplayDimension)mBaseObject.AddChamferDim(x, y, z)).CreateOrNull();

        /// <summary>
        /// 
        /// </summary>
        /// <param name="x">X dimension</param>
        /// <param name="y">Y dimension</param>
        /// <param name="z">Z dimension</param>
        /// <returns>The hole cutout <see cref="ModelDisplayDimension"/> if successful. Null if not.</returns>
        /// <remarks>Make sure to select the hole sketch circle before running this command</remarks>
        public ModelDisplayDimension AddHoleCutout(double x, double y, double z)
            => new ModelDisplayDimension((IDisplayDimension)mBaseObject.AddHoleCallout2(x, y, z)).CreateOrNull();

        /// <summary>
        /// Re-aligns the selected ordinate dimension if it was previously broken
        /// </summary>
        public void AlignOrdinateDimension() => mBaseObject.AlignOrdinate();

        #endregion

        #region Annotation Methods

        /// <summary>
        /// Attaches an existing annotation to a drawing sheet or view
        /// </summary>
        /// <param name="option">The attach option</param>
        /// <returns>True if successful, false if not</returns>
        /// <remarks>
        ///     Remember to select the annotation and if attaching to a view select an
        ///     element on the view also before running this command
        /// </remarks>
        public bool AttachAnnotation(AttachAnnotationOption option) => mBaseObject.AttachAnnotation((int)option);

        /// <summary>
        /// Attempts to attach unattached dimensions, for example in an imported DXF file
        /// </summary>
        public void AttachDimensions() => mBaseObject.AttachDimensions();

        public ITableAnnotation InsertTable(
            swBOMConfigurationAnchorType_e anchor,
            XYZ position,
            int rowCount,
            int columnCount,
            string templatePath = "")
        {
            return mBaseObject.InsertTableAnnotation2(
                // True to anchor the table to the general table
                // anchor point and ignore any coordinates specified
                // for X and Y, or false to use the coordinates specified
                // for X and Y
                false,
                // X coordinate to insert this table annotation
                position.X,
                // Y coordinate to insert this table annotation
                position.Y,
                // Type of anchor as defined in swBOMConfigurationAnchorType_e (see Remarks)
                (int)anchor,
                // Path and filename of the general table template to use  (see Remarks)
                templatePath,
                // Number of rows in the table annotation
                rowCount,
                // Number of columns in the table annotation
                columnCount);

            // Remarks
            // If TableTemplate is... Then..

            // A valid path and filename
            // AnchorType and Columns are ignored, and the information
            // from the table template is used instead

            // Empty
            // General table is inserted based only on the other input arguments
        }

        #endregion

        #region Line Style Methods

        /// <summary>
        /// Adds a line style to the drawing document
        /// </summary>
        /// <param name="styleName">The name of the style</param>
        /// <param name="boldLineEnds">True to have bold dots at each end of the line</param>
        /// <param name="segments">Segments. Positive numbers are dashes, negative are gaps</param>
        /// <returns></returns>
        /// <example>
        ///     AddLineStyle("NewStyle", true, 1.25,-0.5,0.5,-0.5);
        ///     To add a new line like this:
        ///     -----  --  -----  --  -----  --
        /// </example>
        public bool AddLineStyle(string styleName, bool boldLineEnds, params double[] segments)
        {
            // Set line end style
            var segmentString = boldLineEnds ? "B," : "A,";

            // Add segments
            segmentString += string.Join(",", segments.Select(f => f.ToString()));

            // Add line style
            return mBaseObject.AddLineStyle(styleName, segmentString);
        }

        #endregion
    }
}

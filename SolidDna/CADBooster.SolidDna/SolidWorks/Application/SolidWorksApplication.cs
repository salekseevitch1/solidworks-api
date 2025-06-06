﻿using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace CADBooster.SolidDna
{
    /// <summary>
    /// Represents the current SolidWorks application
    /// </summary>
    public partial class SolidWorksApplication : SharedSolidDnaObject<SldWorks>
    {
        #region Protected Members

        /// <summary>
        /// The cookie to the current instance of SolidWorks we are running in.
        /// </summary>
        protected int mSwCookie;

        /// <summary>
        /// The currently active document
        /// </summary>
        protected Model mActiveModel;

        #endregion

        #region Private Members

        /// <summary>
        /// SolidWorks does not fire <see cref="FileOpenPostNotify"/> for view-only ('Large Design Review' or 'Quick-view') models, so we fire the event ourselves.
        /// Skip firing <see cref="FileOpenPostNotify"/> if the model path is already in this list. File path is removed when file is closed.
        /// </summary>
        private static readonly List<string> ViewOnlyFilePaths = new List<string>();

        /// <summary>
        /// Locking object for synchronizing the disposing of SolidWorks and reloading active model info.
        /// </summary>
        private readonly object mDisposingLock = new object();

        #endregion

        #region Public Properties

        /// <summary>
        /// The currently active model
        /// </summary>
        public Model ActiveModel => mActiveModel;

        /// <summary>
        /// The type of SolidWorks application that is currently running.
        /// </summary>
        public SolidWorksApplicationType ApplicationType { get; }

        /// <summary>
        /// The command manager
        /// </summary>
        public CommandManager CommandManager { get; }

        /// <summary>
        /// Whether we are currently connected to 3DExperience. Was added in SolidWorks 2022, so we always return None for older versions.
        /// </summary>
        public ConnectionStatus3DExperience ConnectionStatus3DExperience => SolidWorksVersion.Version < 2022
                ? ConnectionStatus3DExperience.None
                : (ConnectionStatus3DExperience)BaseObject.Get3DExperienceState();

        /// <summary>
        /// True if the application is disposing
        /// </summary>
        public bool Disposing { get; private set; }

        /// <summary>
        /// Various preferences for SolidWorks
        /// </summary>
        public SolidWorksPreferences Preferences { get; }

        /// <summary>
        /// Gets the current SolidWorks version information
        /// </summary>
        public SolidWorksVersion SolidWorksVersion { get; }

        /// <summary>
        /// The SolidWorks instance cookie
        /// </summary>
        public int SolidWorksCookie => mSwCookie;

        /// <summary>
        /// The SolidWorks Math Utility object
        /// </summary>
        public MathUtility MathUtility => BaseObject.GetMathUtility() as MathUtility;

        #endregion

        #region Public Events

        /// <summary>
        /// Called when the currently active file has been saved
        /// </summary>
        public event Action<string, Model> ActiveFileSaved = (path, model) => { };

        /// <summary>
        /// Called when any information about the currently active model has changed
        /// </summary>
        public event Action<Model> ActiveModelInformationChanged = (model) => { };

        /// <summary>
        /// Called when a new file has been created
        /// </summary>
        public event Action<Model> FileCreated = (model) => { };

        /// <summary>
        /// Called when a file has been opened
        /// </summary>
        public event Action<string, Model> FileOpened = (path, model) => { };

        /// <summary>
        /// Called when SolidWorks is idle
        /// </summary>
        public event Action Idle = () => { };

        /// <summary>
        /// Called when SolidWorks is about to close.
        /// </summary>
        public event Action SolidWorksClosing = () => { };

        #endregion

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public SolidWorksApplication(SldWorks solidWorks, int cookie) : base(solidWorks)
        {
            // Set properties that never change
            Preferences = new SolidWorksPreferences();
            SolidWorksVersion = GetSolidWorksVersion();
            ApplicationType = GetApplicationType(); // do this after setting the SolidWorks version.

            // Store cookie ID
            mSwCookie = cookie;

            // Hook into main events
            BaseObject.ActiveModelDocChangeNotify += ActiveModelChanged;
            BaseObject.DestroyNotify += OnSolidWorksClosing;
            BaseObject.FileOpenPostNotify += FileOpenPostNotify;
            BaseObject.FileNewNotify2 += FileNewPostNotify;
            BaseObject.OnIdleNotify += OnIdleNotify;

            // If we have a cookie...
            if (cookie > 0)
                // Get command manager
                CommandManager = new CommandManager(UnsafeObject.GetCommandManager(mSwCookie));

            // Get whatever the current model is on load
            ReloadActiveModelInformation();
        }

        #endregion

        #region Public Callback Events

        /// <summary>
        /// Informs this class that the active model may have changed and it should be reloaded
        /// </summary>
        public void RequestActiveModelChanged()
        {
            ReloadActiveModelInformation();
        }

        #endregion

        #region Application properties

        /// <summary>
        /// Get the type of SolidWorks application we are running in.
        /// Options: Desktop, 3DEXPERIENCE SolidWorks (= SolidWorks Connected) or Desktop with the 3DEXPERIENCE connector add-in.
        /// API was added in SolidWorks 2020, so we always return Desktop for older versions.
        /// </summary>
        /// <returns></returns>
        private SolidWorksApplicationType GetApplicationType()
        {
            // If we are running in a version older than 2020, we are always running in the desktop application.
            return SolidWorksVersion.Version < 2020 ? SolidWorksApplicationType.Desktop : (SolidWorksApplicationType)BaseObject.ApplicationType;
        }

        /// <summary>
        /// Gets the current SolidWorks version information
        /// </summary>
        /// <returns></returns>
        protected SolidWorksVersion GetSolidWorksVersion()
        {
            // Wrap any error
            return SolidDnaErrors.Wrap(() =>
            {
                // Get version string (such as 32.2.0 for 2024 SP2.0)
                var revisionNumber = BaseObject.RevisionNumber();

                // Get revision string (such as sw2024_SP20)
                // Get build number (such as d240722.003) 
                // Get the hot fix string (often empty)
                BaseObject.GetBuildNumbers2(out var revisionString, out var buildNumber, out var hotfixString);

                return new SolidWorksVersion(revisionNumber, revisionString, buildNumber, hotfixString);
            },
                SolidDnaErrorTypeCode.SolidWorksApplication,
                SolidDnaErrorCode.SolidWorksApplicationVersionError);
        }

        #endregion

        #region SolidWorks Event Methods

        /// <summary>
        /// Called when SolidWorks is idle
        /// </summary>
        /// <returns></returns>
        private int OnIdleNotify()
        {
            // Wrap any error
            SolidDnaErrors.Wrap(() =>
            {
                // Inform listeners
                Idle();
            },
                SolidDnaErrorTypeCode.SolidWorksApplication,
                SolidDnaErrorCode.SolidWorksApplicationIdleNotificationError);

            // NOTE: 0 is OK, anything else is an error
            return 0;
        }

        /// <summary>
        /// Called when SolidWorks is about to close.
        /// </summary>
        /// <returns></returns>
        private int OnSolidWorksClosing()
        {
            // Inform listeners
            SolidWorksClosing();

            // NOTE: 0 is OK, anything else is an error
            return 0;
        }

        #region File New

        /// <summary>
        /// Called after a new file has been created.
        /// <see cref="ActiveModel"/> is updated to the new file before this event is called.
        /// </summary>
        /// <param name="newDocument"></param>
        /// <param name="documentType"></param>
        /// <param name="templatePath"></param>
        /// <returns></returns>
        private int FileNewPostNotify(object newDocument, int documentType, string templatePath)
        {
            // Inform listeners
            FileCreated(mActiveModel);

            // IMPORTANT: This is needed after a new file is created as the model COM reference
            //            is created on ActiveModelChanged, and then the file is created after
            // 
            //            This gives a COM reference that fires the FileSaveAsPreNotify event
            //            but then gets disposed and we no longer have any hooks to the active
            //            file so no further events of file save or anything to do with the 
            //            active model fire.
            //
            //            Reloading them at this moment fixes that issue. Then the next issue
            //            is that after the model FileSavePostNotify is fired, it will dispose
            //            of its COM reference again if this is the first time the file is 
            //            saved. To fix that we wait for idle and reload the model information
            //            again. This fix is inside Model.cs FileSavePostNotify
            ReloadActiveModelInformation();

            // NOTE: 0 is OK, anything else is an error
            return 0;
        }

        #endregion

        #region File Open

        /// <summary>
        /// Called after a file has finished opening
        /// </summary>
        /// <param name="filePath">The path to the file being opened</param>
        /// <returns></returns>
        private int FileOpenPostNotify(string filePath)
        {
            // Wrap any error
            SolidDnaErrors.Wrap(() =>
            {
                // And update all properties and models
                ReloadActiveModelInformation();

                // Inform listeners
                FileOpened(filePath, mActiveModel);
            },
                SolidDnaErrorTypeCode.SolidWorksApplication,
                SolidDnaErrorCode.SolidWorksApplicationFilePostOpenError);

            // NOTE: 0 is OK, anything else is an error
            return 0;
        }

        #endregion

        #region Model Changed

        /// <summary>
        /// Called when the active model has changed
        /// </summary>
        /// <returns></returns>
        private int ActiveModelChanged()
        {
            // Wrap any error
            SolidDnaErrors.Wrap(() =>
            {
                // View Only mode (Large Assembly Review and Quick View) does not fire the FileOpenPostNotify event, so we catch these models here.
                var activeDoc = BaseObject.IActiveDoc2;
                if (activeDoc != null)
                {
                    var loadingInViewOnlyMode = activeDoc.IsOpenedViewOnly();
                    var path = activeDoc.GetPathName().ToLower();
                    if (loadingInViewOnlyMode && !ViewOnlyFilePaths.Contains(path))
                    {
                        FileOpenPostNotify(path);
                        ViewOnlyFilePaths.Add(path);
                        return;
                    }
                }

                // If we got here, it isn't the current document so reload the data
                ReloadActiveModelInformation();
            },
                SolidDnaErrorTypeCode.SolidWorksApplication,
                SolidDnaErrorCode.SolidWorksApplicationActiveModelChangedError);

            // NOTE: 0 is OK, anything else is an error
            return 0;
        }

        #endregion

        #endregion

        #region Active Model

        /// <summary>
        /// Remove filepath from ViewOnlyFilePaths list after closing a 'View-only'-model.
        /// </summary>
        /// <param name="filePath"></param>
        internal static void RemoveViewOnlyFilePath(string filePath) => ViewOnlyFilePaths.Remove(filePath?.ToLower());

        /// <summary>
        /// Reloads all the variables, data and COM objects for the newly available SolidWorks model/state
        /// </summary>
        private void ReloadActiveModelInformation()
        {
            // First clean-up any previous SW data
            CleanActiveModelData();

            // Now get the new data
            mActiveModel = BaseObject.IActiveDoc2 == null || BaseObject.GetDocumentCount() == 0 ? null : new Model(BaseObject.IActiveDoc2);

            // Listen out for events
            if (mActiveModel != null)
            {
                mActiveModel.ModelSaved += ActiveModel_Saved;
                mActiveModel.ModelInformationChanged += ActiveModel_InformationChanged;
                mActiveModel.ModelClosing += ActiveModel_Closing;
            }

            // Inform listeners
            ActiveModelInformationChanged(mActiveModel);
        }

        /// <summary>
        /// Disposes of any active model-specific data ready for refreshing
        /// </summary>
        private void CleanActiveModelData()
        {
            // Active model
            mActiveModel?.Dispose();
        }

        #region Event Callbacks

        /// <summary>
        /// Called when the active model has informed us its information has changed
        /// </summary>
        private void ActiveModel_InformationChanged()
        {
            // Inform listeners
            ActiveModelInformationChanged(mActiveModel);
        }

        /// <summary>
        /// Called when the active document is closed
        /// </summary>
        private void ActiveModel_Closing()
        {
            // 
            // NOTE: There is no event to detect when all documents are closed 
            // 
            //       So, each model that is closing (not closed) wait 200ms 
            //       then check on the current number of active documents
            //       or if ActiveDoc is already set to null.
            //
            //       If ActiveDoc is null or the document count is 0 at that 
            //       moment in time, do an active model information refresh.
            //
            //       If another document opens in the meantime it won't fire
            //       but that's fine as the active doc changed event will fire
            //       in that case anyway
            //

            // Check for every file if it may have been the last one.
            Task.Run(async () =>
            {
                // Wait for it to close
                await Task.Delay(200);

                // Lock to prevent Disposing to change while this section is running.
                lock (mDisposingLock)
                {
                    if (Disposing)
                        // If we are disposing SolidWorks, there is no need to reload active model info.
                        return;

                    // Now if we have none open, reload information
                    // ActiveDoc is quickly set to null after the last document is closed
                    // GetDocumentCount takes longer to go to zero for big assemblies, but it might be a more reliable indicator.
                    if (BaseObject?.ActiveDoc == null || BaseObject?.GetDocumentCount() == 0)
                        ReloadActiveModelInformation();

                }
            });
        }

        /// <summary>
        /// Called when the currently active file has been saved
        /// </summary>
        private void ActiveModel_Saved()
        {
            // Inform listeners
            ActiveFileSaved(mActiveModel?.FilePath, mActiveModel);
        }

        #endregion

        #endregion

        #region Create a new file

        /// <summary>
        /// Create a new assembly. Throws if it fails.
        /// </summary>
        /// <param name="templatePath">Your preferred assembly template path. Pass null to use the default assembly template.</param>
        /// <returns></returns>
        public Model CreateAssembly(string templatePath = null)
        {
            // If the user did not pass a template path, we get the default template path from SolidWorks.
            if (templatePath.IsNullOrEmpty())
                templatePath = Preferences.DefaultAssemblyTemplate;

            return CreateFile(templatePath);
        }

        /// <summary>
        /// Create a new drawing with a standard paper size. Throws if it fails.
        /// </summary>
        /// <param name="paperSize"></param>
        /// <param name="templatePath">Your preferred drawing template path. Pass null to use the default drawing template.</param>
        /// <returns></returns>
        public Model CreateDrawing(swDwgPaperSizes_e paperSize, string templatePath = null)
        {
            // If the user did not pass a template path, we get the default template path from SolidWorks.
            if (templatePath.IsNullOrEmpty())
                templatePath = Preferences.DefaultDrawingTemplate;

            return CreateFile(templatePath, paperSize);
        }

        /// <summary>
        /// Create a new drawing with a custom paper size. Throws if it fails.
        /// </summary>
        /// <param name="height"></param>
        /// <param name="width"></param>
        /// <param name="templatePath">Your preferred drawing template path. Pass null to use the default drawing template.</param>
        /// <returns></returns>
        public Model CreateDrawing(double width, double height, string templatePath = null)
        {
            // If the user did not pass a template path, we get the default template path from SolidWorks.
            if (templatePath.IsNullOrEmpty())
                templatePath = Preferences.DefaultDrawingTemplate;

            return CreateFile(templatePath, swDwgPaperSizes_e.swDwgPapersUserDefined, width, height);
        }

        /// <summary>
        /// Create a new part. Throws if it fails.
        /// </summary>
        /// <param name="templatePath">Your preferred part template path. Pass null to use the default part template.</param>
        /// <returns></returns>
        public Model CreatePart(string templatePath = null)
        {
            // If the user did not pass a template path, we get the default template path from SolidWorks.
            if (templatePath.IsNullOrEmpty())
                templatePath = Preferences.DefaultPartTemplate;

            return CreateFile(templatePath);
        }

        /// <summary>
        /// Create a new model. Throws if it fails.
        /// </summary>
        /// <param name="templatePath"></param>
        /// <param name="paperSize"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <returns></returns>
        private Model CreateFile(string templatePath, swDwgPaperSizes_e paperSize = swDwgPaperSizes_e.swDwgPaperA3size, double width = 0, double height = 0)
        {
            // Wrap any error
            return SolidDnaErrors.Wrap(() =>
            {
                // Create the new file
                var swModel = UnsafeObject.INewDocument2(templatePath, (int)paperSize, width, height);

                // If the modelDoc is null, creating a new file failed
                if (swModel == null)
                    throw new Exception("Failed to create a new file");

                // If we have a value, we wrap it in a Model
                return new Model(swModel);
            }, SolidDnaErrorTypeCode.File, SolidDnaErrorCode.FileCreateError);
        }

        #endregion

        #region Open/Close Models

        /// <summary>
        /// Loops all open documents returning a safe <see cref="Model"/> for each document,
        /// disposing of the COM reference after its use
        /// </summary>
        /// <returns></returns>
        public IEnumerable<Model> OpenDocuments()
        {
            // Loop each child
            foreach (ModelDoc2 modelDoc in (object[])BaseObject.GetDocuments())
            {
                // Create safe model
                using (var model = new Model(modelDoc))
                    // Return it
                    yield return model;
            }
        }

        /// <summary>
        /// Open a part, assembly or drawing by its file path.
        /// Use <see cref="OpenFile(IDocumentSpecification)"/> to open a model with extra options.
        /// </summary>
        /// <param name="filePath">The path to the file</param>
        /// <param name="options">The options to use when opening the file (flags, so use pipes | to combine options)</param>
        /// <param name="configuration">The name of the configuration you want to open. If you skip this parameter, SolidWorks will open the configuration is which the model was last saved.</param>
        public Model OpenFile(string filePath, OpenDocumentOptions options = OpenDocumentOptions.None, string configuration = null)
        {
            // Wrap any error
            return SolidDnaErrors.Wrap(() =>
            {
                // Get file type. Throws when the string is not long enough, but that gets caught by the SolidDnaErrors.Wrap method anyway.
                var fileExtension = filePath.Substring(filePath.Length - 7);
                var fileType =
                    fileExtension.Equals(".sldprt", StringComparison.OrdinalIgnoreCase) ? DocumentType.Part :
                    fileExtension.Equals(".sldasm", StringComparison.OrdinalIgnoreCase) ? DocumentType.Assembly :
                    fileExtension.Equals(".slddrw", StringComparison.OrdinalIgnoreCase) ? DocumentType.Drawing : throw new ArgumentException("Unknown file type");

                // Set errors and warnings
                var errors = 0;
                var warnings = 0;

                // Attempt to open the document
                var swModel = BaseObject.OpenDoc6(filePath, (int)fileType, (int)options, configuration, ref errors, ref warnings);

                // TODO: Read errors into enums for better reporting
                // For now just check if model is not null
                if (swModel == null)
                    throw new Exception($"Failed to open file. Errors {errors}, Warnings {warnings}");

                // Return new model
                return new Model(swModel);
            },
                SolidDnaErrorTypeCode.SolidWorksApplication,
                SolidDnaErrorCode.SolidWorksModelOpenFileError);
        }

        /// <summary>
        /// Open a part, assembly or drawing and fully control how the file is opened.
        /// Can be used to open local files and files from 3DExperience.
        /// </summary>
        /// <param name="documentSpecification">File properties and instructions on how to open the file.
        /// Properties that are not valid when opening a file from 3DExperience: FileName, ConfigurationName,
        /// DocumentType, DisplayState, InteractiveAdvancedOpen, InteractiveComponentSelection, LoadExternalReferencesInMemory.</param>
        /// <returns></returns>
        public Model OpenFile(IDocumentSpecification documentSpecification)
        {
            // Wrap any error
            return SolidDnaErrors.Wrap(() =>
                {
                    // Attempt to open the document
                    var swModel = BaseObject.OpenDoc7(documentSpecification);

                    // TODO: Read errors into enums for better reporting
                    // For now just check if model is not null
                    if (swModel == null)
                        throw new Exception($"Failed to open file. Errors {documentSpecification.Error}, Warnings {documentSpecification.Warning}");

                    // Return new model
                    return new Model(swModel);
                },
                SolidDnaErrorTypeCode.SolidWorksApplication,
                SolidDnaErrorCode.SolidWorksModelOpenFileError);
        }

        /// <summary>
        /// Open a part, assembly or drawing by its PLM ID (its unique 3DExperience ID).
        /// 3DExperience treats configurations as unique models, so you cannot specify a configuration when opening a model by its PLM ID.
        /// Use <see cref="OpenFile(IDocumentSpecification)"/> to open a model with extra options.
        /// </summary>
        /// <param name="plmId">The unique ID of this model. Consists of numbers and letters, seems to be 32 characters long.</param>
        /// <returns></returns>
        public Model OpenFileFrom3DExperience(string plmId)
        {
            // Get a new object specification
            var swDocSpecification = (IDocumentSpecification)BaseObject.GetOpenDocSpec("");

            // Get a detailed 3DX object specification
            var plmObjectSpecification = (IPLMObjectSpecification)swDocSpecification.PLMObjectSpecification;

            // Set the PLM ID
            plmObjectSpecification.PLMID = plmId;

            // Call the Open method that takes a document specification parameter
            return OpenFile(swDocSpecification);
        }

        /// <summary>
        /// Closes a file
        /// </summary>
        /// <param name="filePath">The path to the file</param>
        public void CloseFile(string filePath)
        {
            // Wrap any error
            SolidDnaErrors.Wrap(() =>
            {
                BaseObject.CloseDoc(filePath);
            },
                SolidDnaErrorTypeCode.SolidWorksApplication,
                SolidDnaErrorCode.SolidWorksModelCloseFileError);
        }

        #endregion

        #region Save Data

        /// <summary>
        /// Gets an <see cref="IExportPdfData"/> object for use with a <see cref="PdfExportData"/>
        /// object used in <see cref="Model.SaveAs(string, SaveAsVersion, SaveAsOptions, PdfExportData)"/> call
        /// </summary>
        /// <returns></returns>
        public IExportPdfData GetPdfExportData()
        {
            // NOTE: No point making our own enumerator for the export file type
            //       as right now and for many years it's only ever been
            //       1 for PDF. I do not see this ever changing.
            return BaseObject.GetExportFileData((int)swExportDataFileType_e.swExportPdfData) as IExportPdfData;
        }

        #endregion

        #region Materials

        /// <summary>
        /// Gets a list of all materials in SolidWorks
        /// </summary>
        /// <param name="databasePath">If specified, limits the results to the specified database full path</param>
        public List<Material> GetMaterials(string databasePath = null)
        {
            // Wrap any error
            return SolidDnaErrors.Wrap(() =>
            {
                // Create an empty list
                var list = new List<Material>();

                // If we are using a specified database, use that
                if (databasePath != null)
                    ReadMaterials(databasePath, ref list);
                else
                {
                    // Otherwise, get all known ones
                    // Get the list of material databases (full paths to SLDMAT files)
                    var databasePaths = (string[])BaseObject.GetMaterialDatabases();

                    // Get materials from each
                    if (databasePaths != null)
                        foreach (var path in databasePaths)
                            ReadMaterials(path, ref list);
                }

                // Order the list
                return list.OrderBy(f => f.DisplayName).ToList();
            },
                SolidDnaErrorTypeCode.SolidWorksApplication,
                SolidDnaErrorCode.SolidWorksApplicationGetMaterialsError);
        }

        /// <summary>
        /// Attempts to find the material from a SolidWorks material database file (SLDMAT)
        /// If found, returns the full information about the material
        /// </summary>
        /// <param name="databasePath">The full path to the database</param>
        /// <param name="materialName">The material name to find</param>
        /// <returns></returns>
        public Material FindMaterial(string databasePath, string materialName)
        {
            // Wrap any error
            return SolidDnaErrors.Wrap(() =>
            {
                // Get all materials from the database
                var materials = GetMaterials(databasePath);

                // Return if found the material with the same name
                return materials?.FirstOrDefault(f => string.Equals(f.Name, materialName, StringComparison.InvariantCultureIgnoreCase));
            },
                SolidDnaErrorTypeCode.SolidWorksApplication,
                SolidDnaErrorCode.SolidWorksApplicationFindMaterialsError,
                nameof(SolidDnaErrorCode.SolidWorksApplicationFindMaterialsError));
        }

        #region Private Helpers

        /// <summary>
        /// Reads the material database and adds the materials to the given list
        /// </summary>
        /// <param name="databasePath">The database to read</param>
        /// <param name="list">The list to add materials to</param>
        private static void ReadMaterials(string databasePath, ref List<Material> list)
        {
            // First make sure the file exists
            if (!File.Exists(databasePath))
                throw new SolidDnaException(
                    SolidDnaErrors.CreateError(
                        SolidDnaErrorTypeCode.SolidWorksApplication,
                        SolidDnaErrorCode.SolidWorksApplicationGetMaterialsFileNotFoundError));

            try
            {
                // File should be an XML document, so attempt to read that
                using (var stream = File.Open(databasePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    // Try and parse the Xml
                    var xmlDoc = XDocument.Load(stream);

                    var materials = new List<Material>();

                    // Iterate all classification nodes and inside are the materials
                    xmlDoc.Root.Elements("classification").ToList().ForEach(f =>
                    {
                        // Get classification name
                        var classification = f.Attribute("name")?.Value;

                        // Iterate all materials
                        f.Elements("material").ToList().ForEach(material =>
                        {
                            // Add them to the list
                            materials.Add(new Material
                            {
                                DatabasePathOrFilename = databasePath,
                                DatabaseFileFound = true,
                                Classification = classification,
                                Name = material.Attribute("name")?.Value,
                                Description = material.Attribute("description")?.Value,
                            });
                        });
                    });

                    // If we found any materials, add them
                    if (materials.Count > 0)
                        list.AddRange(materials);
                }
            }
            catch (Exception ex)
            {
                // If we crashed for any reason during parsing, wrap in SolidDna exception
                if (!File.Exists(databasePath))
                    throw new SolidDnaException(
                        SolidDnaErrors.CreateError(
                            SolidDnaErrorTypeCode.SolidWorksApplication,
                            SolidDnaErrorCode.SolidWorksApplicationGetMaterialsFileFormatError),
                            ex);
            }
        }

        #endregion

        #endregion

        #region Preferences

        /// <summary>
        /// Gets the specified user preference value
        /// </summary>
        /// <param name="preference">The preference to get</param>
        /// <returns></returns>
        public double GetUserPreferencesDouble(swUserPreferenceDoubleValue_e preference) => BaseObject.GetUserPreferenceDoubleValue((int)preference);

        /// <summary>
        /// Sets the specified user preference value
        /// </summary>
        /// <param name="preference">The preference to set</param>
        /// <param name="value">The new value of the preference</param>
        /// <returns></returns>
        public bool SetUserPreferencesDouble(swUserPreferenceDoubleValue_e preference, double value) => BaseObject.SetUserPreferenceDoubleValue((int)preference, value);

        /// <summary>
        /// Gets the specified user preference value
        /// </summary>
        /// <param name="preference">The preference to get</param>
        /// <returns></returns>
        public int GetUserPreferencesInteger(swUserPreferenceIntegerValue_e preference) => BaseObject.GetUserPreferenceIntegerValue((int)preference);

        /// <summary>
        /// Sets the specified user preference value
        /// </summary>
        /// <param name="preference">The preference to set</param>
        /// <param name="value">The new value of the preference</param>
        /// <returns></returns>
        public bool SetUserPreferencesInteger(swUserPreferenceIntegerValue_e preference, int value) => BaseObject.SetUserPreferenceIntegerValue((int)preference, value);

        /// <summary>
        /// Gets the specified user preference value
        /// </summary>
        /// <param name="preference">The preference to get</param>
        /// <returns></returns>
        public string GetUserPreferencesString(swUserPreferenceStringValue_e preference) => BaseObject.GetUserPreferenceStringValue((int)preference);

        /// <summary>
        /// Sets the specified user preference value
        /// </summary>
        /// <param name="preference">The preference to set</param>
        /// <param name="value">The new value of the preference</param>
        /// <returns></returns>
        public bool SetUserPreferencesString(swUserPreferenceStringValue_e preference, string value) => BaseObject.SetUserPreferenceStringValue((int)preference, value);

        /// <summary>
        /// Gets the specified user preference value
        /// </summary>
        /// <param name="preference">The preference to get</param>
        /// <returns></returns>
        public bool GetUserPreferencesToggle(swUserPreferenceToggle_e preference) => BaseObject.GetUserPreferenceToggle((int)preference);

        /// <summary>
        /// Sets the specified user preference value
        /// </summary>
        /// <param name="preference">The preference to set</param>
        /// <param name="value">The new value of the preference</param>
        /// <returns></returns>
        public void SetUserPreferencesToggle(swUserPreferenceToggle_e preference, bool value) => BaseObject.SetUserPreferenceToggle((int)preference, value);

        #endregion

        #region Taskpane Methods

        /// <summary>
        /// Attempts to create a task pane. Uses a single icon.
        /// </summary>
        /// <param name="iconPath">
        /// An absolute path to an icon to use for the taskpane.
        /// The bitmap should be 16 colors and 16 x 18 (width x height) pixels. 
        /// Any portions of the bitmap that are white (RGB 255,255,255) will be transparent.
        /// </param>
        /// <param name="toolTip">The title text to show at the top of the taskpane</param>
        public async Task<Taskpane> CreateTaskpaneAsync(string iconPath, string toolTip)
        {
            // Wrap any error creating the taskpane in a SolidDna exception
            return SolidDnaErrors.Wrap<Taskpane>(() =>
            {
                // Attempt to create the taskpane
                var comTaskpane = BaseObject.CreateTaskpaneView2(iconPath, toolTip);

                // If we fail, return null
                if (comTaskpane == null)
                    return null;

                // If we succeed, create SolidDna object
                return new Taskpane(comTaskpane);
            },
                SolidDnaErrorTypeCode.SolidWorksTaskpane,
                SolidDnaErrorCode.SolidWorksTaskpaneCreateError);
        }

        /// <summary>
        /// Attempts to create a task pane. Uses a list of PNG icon sizes: 20, 32, 40, 64, 96 and 128 pixels square.
        /// </summary>
        /// <param name="iconPathFormat">The absolute path to all icons, based on a string format of the absolute path. Replaces "{0}" in the string with the icons sizes. 
        /// For example C:\Folder\icons{0}.png
        /// </param>
        /// <param name="toolTip">The title text to show at the top of the taskpane</param>
        public async Task<Taskpane> CreateTaskpaneAsync2(string iconPathFormat, string toolTip)
        {
            // Wrap any error creating the taskpane in a SolidDna exception
            return SolidDnaErrors.Wrap<Taskpane>(() =>
                {
                    // Get up to six icon paths
                    var icons = Icons.GetPathArrayFromPathFormat(iconPathFormat);

                    // Attempt to create the taskpane
                    var comTaskpane = BaseObject.CreateTaskpaneView3(icons, toolTip);

                    // If we fail, return null
                    if (comTaskpane == null)
                        return null;

                    // If we succeed, create SolidDna object
                    return new Taskpane(comTaskpane);
                },
                SolidDnaErrorTypeCode.SolidWorksTaskpane,
                SolidDnaErrorCode.SolidWorksTaskpaneCreateError);
        }

        #endregion

        #region User Interaction

        /// <summary>
        /// Pops up a message box to the user with the given message
        /// </summary>
        /// <param name="message">The message to display to the user</param>
        /// <param name="icon">The severity icon to display</param>
        /// <param name="buttons">The buttons to display</param>
        public SolidWorksMessageBoxResult ShowMessageBox(string message, SolidWorksMessageBoxIcon icon = SolidWorksMessageBoxIcon.Information, SolidWorksMessageBoxButtons buttons = SolidWorksMessageBoxButtons.Ok)
        {
            // Send message to user
            return (SolidWorksMessageBoxResult)BaseObject.SendMsgToUser2(message, (int)icon, (int)buttons);
        }

        /// <summary>
        /// Activates the specified model within the given part document.
        /// </summary>
        /// <remarks>This method activates the model specified by the <paramref name="partDocument"/> and
        /// optionally rebuilds it based on the <paramref name="rebuildOnActivation"/> parameter. The activation result
        /// is provided via the <paramref name="error"/> output parameter.</remarks>
        /// <param name="partDocument">The part document containing the model to activate. Cannot be null.</param>
        /// <param name="rebuildOnActivation">Specifies whether the model should be rebuilt upon activation. Use <see cref="swRebuildOnActivation_e"> to
        /// indicate the desired rebuild behavior.</param>
        /// <param name="error">When the method returns, contains the result of the activation operation. Use <see
        /// cref="swActivateDocError_e"> to determine the success or failure of the activation.</param>
        /// <returns>The activated model if the operation succeeds; otherwise, <see langword="null"/>.</returns>
        public Model? ActivateModel(
            PartDocument partDocument,
            swRebuildOnActivation_e rebuildOnActivation = swRebuildOnActivation_e.swRebuildActiveDoc)
            => ActivateModel(partDocument.Title, rebuildOnActivation, out _);

        /// <summary>
        /// Activates the specified model within the drawing document.
        /// </summary>
        /// <remarks>Use this method to activate a model within a drawing document, optionally rebuilding
        /// it during activation. The method returns <see langword="null"/> if the activation fails, and the <paramref
        /// name="error"/> parameter provides additional information about the failure.</remarks>
        /// <param name="drawingDocument">The drawing document containing the model to activate. Cannot be null.</param>
        /// <param name="rebuildOnActivation">Specifies whether the model should be rebuilt upon activation.</param>
        /// <param name="error">When the method returns, contains the result of the activation operation.  This parameter is passed
        /// uninitialized.</param>
        /// <returns>The activated model if the operation succeeds; otherwise, <see langword="null"/>.</returns>
        public Model? ActivateModel(
            DrawingDocument drawingDocument,
            swRebuildOnActivation_e rebuildOnActivation = swRebuildOnActivation_e.swRebuildActiveDoc) 
            => ActivateModel(drawingDocument.Title, rebuildOnActivation, out _);

        /// <summary>
        /// Activates a model in the application by its title and returns the activated model.
        /// </summary>
        /// <remarks>This method attempts to activate a model by its title. If the activation is
        /// successful, the method returns a <see cref="Model"/> object. If the activation fails, the method returns
        /// <see langword="null"/> and sets the <paramref name="error"/> parameter to indicate the failure
        /// reason.</remarks>
        /// <param name="modelTitle">The title of the model to activate. This must match the name of an existing model.</param>
        /// <param name="rebuildOnActivation">Specifies whether the model should be rebuilt upon activation. Use values from <see
        /// cref="swRebuildOnActivation_e"/>.</param>
        /// <param name="error">When the method returns, contains the result of the activation operation. Use values from <see
        /// cref="swActivateDocError_e"/> to determine the outcome.</param>
        /// <returns>A <see cref="Model"/> object representing the activated model, or <see langword="null"/> if the activation
        /// fails.</returns>
        private Model? ActivateModel(
            string modelTitle,
            swRebuildOnActivation_e rebuildOnActivation,
            out swActivateDocError_e error)
        {
            var swError = 0;

            var modelDoc = UnsafeObject.ActivateDoc3(
                modelTitle,
                true,
                (int)rebuildOnActivation,
                ref swError);

            error = (swActivateDocError_e)swError;

            return modelDoc is ModelDoc2 mDoc ? new Model(mDoc) : null;
        }

        #endregion

        #region Dispose

        /// <summary>
        /// Disposing
        /// </summary>
        public override void Dispose()
        {
            lock (mDisposingLock)
            {

                // Flag as disposing
                Disposing = true;

                // Clean active model
                ActiveModel?.Dispose();

                // Dispose command manager
                CommandManager?.Dispose();

                // NOTE: Don't dispose the application, SolidWorks does that itself
                //base.Dispose();
            }
        }

        #endregion
    }
}

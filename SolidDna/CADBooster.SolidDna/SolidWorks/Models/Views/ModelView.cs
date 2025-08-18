using SolidWorks.Interop.sldworks;

namespace CADBooster.SolidDna.Views;

public class ModelView(IModelView comObject) : SolidDnaObject<IModelView>(comObject);
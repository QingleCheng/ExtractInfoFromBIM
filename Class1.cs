using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using View = Autodesk.Revit.DB.View;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI.Selection;
using System.IO;

namespace Project
{
    [Autodesk.Revit.Attributes.Transaction(TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(RegenerationOption.Manual)]



    public class GetAllWindows : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {



            string readme = "elems_info.txt"; 
            string path = "D:/";
            string txtpath = path + "/" + readme;
            DirectoryInfo directoryInfo = new DirectoryInfo(path);
            if (!directoryInfo.Exists)  
            {
                directoryInfo.Create();
            }
            if (!File.Exists(txtpath))
            {
                File.Create(txtpath).Close();
            }
            else
            {
                File.Delete(txtpath);
                File.Create(txtpath).Close();
            }


            UIApplication uiApp = commandData.Application;
            Autodesk.Revit.ApplicationServices.Application app = uiApp.Application;
            Document doc = uiApp.ActiveUIDocument.Document;
            FilteredElementCollector collectorAll = new FilteredElementCollector(doc);// Collect all elements
            collectorAll.OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_Furniture);//Get furniture category

            IList<Element> lists = collectorAll.ToElements();//Convert it to a list of Element 
            String name = "";
            foreach (var item in lists)
            {
                try {
                    StreamWriter sw = new StreamWriter(txtpath, true, System.Text.Encoding.Default);
                    
                    // Extract the level attribute of an element
                    Level level = item.Document.GetElement(item.LevelId) as Level;
                    FamilyInstance familyInstance = item as FamilyInstance; //Extract the room attribute of an element
                    string strParamInfo = null;
                    Parameter param = item.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);//Extract the commemt of an element
                    strParamInfo = param.AsString();
                    
                    string FragilityID = null;
                    Parameter param2 = item.LookupParameter("FragilityID");//Extract the fragility ID of an element
                    FragilityID = param2.AsString();

                    // Get FamilyInstance room name
                    if (familyInstance.Room != null)
                    {
                        name = item.Name + "\t" + level.Name + "\t" + familyInstance.Room.Name+"\t"+ strParamInfo + "\t" + FragilityID;
                        sw.WriteLine(name);
                    }
                    sw.Flush();
                    sw.Close();
                }
                catch (Exception e)
                {

                }

            }
            return Result.Succeeded;
        }

        public void FocusElements(UIApplication uiApp, List<ElementId> elementIds)
        {
            var doc = uiApp.ActiveUIDocument.Document;

            var views = new FilteredElementCollector(doc).OfClass(typeof(View3D));
            if (views.Count() > 0)
            {
                foreach (View item in views)
                {
                    if (item.IsTemplate) continue;
                    uiApp.PostCommand(RevitCommandId.LookupPostableCommandId(PostableCommand.DeactivateView));
                    uiApp.ActiveUIDocument.ActiveView = item;
                    break;
                }
            }
            uiApp.ActiveUIDocument.Selection.SetElementIds(elementIds);
            uiApp.ActiveUIDocument.ShowElements(elementIds);
            uiApp.ActiveUIDocument.RefreshActiveView();
        }
    }
}

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AddLinkFromFile
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            //dialog.InitialDirectory = folder;
            dialog.Multiselect = false;
            dialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            if (dialog.ShowDialog() != DialogResult.OK) return Result.Cancelled;
            string filePath = dialog.FileName;

            string text = System.IO.File.ReadAllText(filePath);
            List<string> textArr = text.Split('\n').ToList();
            textArr = textArr.Select(t => t.Replace("\r", String.Empty)).ToList();

            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            List<ModelPath> listModel = new List<ModelPath>();
            foreach (string modelPath in textArr)
            {
                listModel.Add(ModelPathUtils.ConvertUserVisiblePathToModelPath(modelPath));
            }

            Transaction tr = new Transaction(doc, "Загрузка файлов");
            tr.Start();

            RevitLinkOptions rlo = new RevitLinkOptions(false);

            List<string> successLoadModel = new List<string>();
            List<string> notSuccessLoadModel = new List<string>();
            List<RevitLinkInstance> successLink = new List<RevitLinkInstance>();

            foreach (ModelPath model in listModel)
            {
                try
                {
                    LinkLoadResult linkType = RevitLinkType.Create(doc, model, rlo);
                    RevitLinkInstance instance = RevitLinkInstance.Create(doc, linkType.ElementId);
                    successLoadModel.Add(ModelPathUtils.ConvertModelPathToUserVisiblePath(model));
                    successLink.Add(instance);
                }
                catch (Exception)
                {
                    notSuccessLoadModel.Add(ModelPathUtils.ConvertModelPathToUserVisiblePath(model));
                }
            }

            tr.Commit();
            string messagefinal = null;
            if (successLoadModel.Count > 0)
            {
                messagefinal += "Следующие модели УСПЕШНО загружены: \n";
                successLoadModel.ForEach(t => messagefinal += (t + "\n"));
            }

            if (notSuccessLoadModel.Count > 0)
            {
                messagefinal += "\n";
                messagefinal += "Следующие модели НЕ УДАЛОСЬ загрузить: \n";
                notSuccessLoadModel.ForEach(t => messagefinal += (t + "\n"));
            }
            DialogResult dialogResult;
            if (successLoadModel.Count > 0 || notSuccessLoadModel.Count > 0)
            {
                messagefinal += "\n";
                messagefinal += "Успешно загруженные связи будут выделены, осталось нажать кнопку 'Внедрить связь'";
                
                dialogResult = MessageBox.Show(messagefinal, "Конец");
                uiDoc.Selection.SetElementIds(successLink.Select(x => x.Id).ToList());
                //if (dialogResult == DialogResult.Yes)
                //{
                //    //RevitCommandId id = RevitCommandId.LookupPostableCommandId(PostableCommand.LoadAsGroup);
                //    //uiApp.PostCommand(id);
                //}
            }

            return Result.Succeeded;
        }
    }
}

#region License
/*Данный код опубликован под лицензией Creative Commons Attribution-ShareAlike.
Разрешено использовать, распространять, изменять и брать данный код за основу для производных в коммерческих и
некоммерческих целях, при условии указания авторства и если производные лицензируются на тех же условиях.
Код поставляется "как есть". Автор не несет ответственности за возможные последствия использования.
Зуев Александр, 2020, все права защищены.
This code is listed under the Creative Commons Attribution-ShareAlike license.
You may use, redistribute, remix, tweak, and build upon this work non-commercially and commercially,
as long as you credit the author by linking back and license your new creations under the same terms.
This code is provided 'as is'. Author disclaims any implied warranty.
Zuev Aleksandr, 2020, all rigths reserved.
*/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB; //для работы с элементами модели Revit
using Autodesk.Revit.UI; //для работы с элементами интерфейса
using Autodesk.Revit.UI.Selection; //работы с выделенными элементами

namespace CollapseRebarSchedule
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            ViewSchedule vs = commandData.Application.ActiveUIDocument.ActiveView as ViewSchedule;
            ScheduleDefinition sdef = null;
            if (vs == null)
            {
                Selection sel = commandData.Application.ActiveUIDocument.Selection;
                if (sel.GetElementIds().Count == 0) return Result.Failed;
                ScheduleSheetInstance ssi = doc.GetElement(sel.GetElementIds().First()) as ScheduleSheetInstance;
                if (ssi == null) return Result.Failed;
                if (!ssi.Name.Contains("ВРС")) return Result.Failed;
                vs = doc.GetElement(ssi.ScheduleId) as ViewSchedule;
            }
            sdef = vs.Definition;


            int firstWeightCell = 0;
            int startHiddenFields = 0;
            int borderCell = 9999;

            //определяю первую и последнюю ячейку с массой
            for (int i = 0; i < sdef.GetFieldCount(); i++)
            {
                ScheduleField sfield = sdef.GetField(i);
                string cellName = sfield.GetName();
                if (firstWeightCell == 0)
                {
                    if (char.IsNumber(cellName[0]))
                    {
                        firstWeightCell = i;
                    }
                    else
                    {
                        if (sfield.IsHidden)
                        {
                            startHiddenFields++;
                        }
                    }
                }
                if (cellName.StartsWith("="))
                {
                    borderCell = i;
                    break;
                }
            }

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Отобразить все ячейки");
                for (int i = firstWeightCell; i < borderCell; i++)
                {
                    ScheduleField sfield = sdef.GetField(i);
                    string cellName = sfield.GetName();
                    sfield.IsHidden = false;
                }


                doc.Regenerate();

                TableData tdata = vs.GetTableData();
                TableSectionData tsd = tdata.GetSectionData(SectionType.Body);
                int firstRownumber = tsd.FirstRowNumber;
                int lastRowNumber = tsd.LastRowNumber;
                int rowsCount = lastRowNumber - firstRownumber;

                for (int i = firstWeightCell; i < borderCell; i++)
                {
                    ScheduleField sfield = sdef.GetField(i);
                    string cellName = sfield.GetName();

                    List<string> values = new List<string>();
                    for (int j = firstRownumber; j <= lastRowNumber; j++)
                    {
                        string cellText = tsd.GetCellText(j, i - startHiddenFields);
                        values.Add(cellText);
                    }

                    bool checkOnlyTextAndZeros = OnlyTextAndZeros(values);
                    if (checkOnlyTextAndZeros)
                    {
                        sfield.IsHidden = true;
                    }
                }
                t.Commit();


            }
            return Result.Succeeded;
        }

        private bool OnlyTextAndZeros(List<string> values)
        {
            bool haveZeros = false;
            bool haveNumber = false;
            double val = -1;
            foreach (string s in values)
            {
                val = -1;
                if (string.IsNullOrEmpty(s))
                {
                    continue;
                }
                bool isNumber = double.TryParse(s, out val);
                if (!isNumber)
                {
                    continue;
                }
                if (val > 0)
                {
                    haveNumber = true;
                    continue;
                }
                else
                {
                    haveZeros = true;
                }
            }
            if (haveZeros && !haveNumber)
            {
                return true;
            }
            return false;
        }
    }
}

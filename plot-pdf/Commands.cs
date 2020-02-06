using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.PlottingServices;

using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AcadPlugin
{
    public class Commands
    {
        [CommandMethod("QPLOT")]
        public void Main()
        {

            Document doc = AcAp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            List<ObjectId> ids = null;

            LayoutManager layManager = LayoutManager.Current;
            ObjectId layoutId = layManager.GetLayoutId(layManager.CurrentLayout);
            //Layout layoutObj = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
            Layout layoutObj = layoutId.Open(OpenMode.ForRead) as Layout;
            PlotSettings plotSettingsObj = new PlotSettings(layoutObj.ModelType);
            string setName = "Langamer_Plot";

            using (var trAddPlotSet = db.TransactionManager.StartTransaction())
            {
                // Adiciona o PlotSetup ao DWG atual **********************************************************************************************
                string plotSetDWGName = "Z:\\Lisp\\dwg_plot_setup.dwg";
                Database dbPlotSet = new Database();
                dbPlotSet.ReadDwgFile(plotSetDWGName, FileOpenMode.OpenForReadAndAllShare, true, "");
                using (var trPlotSet = dbPlotSet.TransactionManager.StartTransaction())
                {
                    BlockTable btPlotSet = trPlotSet.GetObject(dbPlotSet.BlockTableId, OpenMode.ForWrite) as BlockTable;
                    var btrPlotSet = trPlotSet.GetObject(btPlotSet[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    Layout loPlotSet = trPlotSet.GetObject(btrPlotSet.LayoutId, OpenMode.ForWrite) as Layout;

                    PlotSettings psPlotSet = new PlotSettings(loPlotSet.ModelType);
                    psPlotSet.CopyFrom(loPlotSet);

                    psPlotSet.AddToPlotSettingsDictionary(db);
                    trPlotSet.Commit();
                    trAddPlotSet.Commit();
                }
            }

            using (var tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
                BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;


                DBDictionary plotSettingsDic = (DBDictionary)tr.GetObject(db.PlotSettingsDictionaryId, OpenMode.ForRead);
                if (!plotSettingsDic.Contains(setName))
                    return;

                ObjectId plotSettingsId = plotSettingsDic.GetAt(setName);

                //layout type
                bool bModel = layoutObj.ModelType;
                plotSettingsObj = tr.GetObject(plotSettingsId, OpenMode.ForRead) as PlotSettings;

                if (plotSettingsObj.ModelType != bModel)
                    return;

                object backgroundPlot = Application.GetSystemVariable("BACKGROUNDPLOT");
                Application.SetSystemVariable("BACKGROUNDPLOT", 0);
                // Fim: Adiciona o PlotSetup ao DWG atual ****************************************************************************************

                // Pega todos os POLYLINE do desenho atual **********************************************************************************************
                // Cast the BlockTableRecord into IEnumerable<T> collection
                IEnumerable<ObjectId> b = btr.Cast<ObjectId>();

                // Using LINQ statement to select LWPOLYLINE from ModelSpace
                ids = (from id1 in b
                       where OpenObject(tr, id1) is BlockReference &&
                             OpenObject(tr, id1).Layer == "LANG-GR-PRANCHA 1"
                       select id1).ToList<ObjectId>();
                // Fim: Pega todos os POLYLINE do desenho atual **********************************************************************************************

                foreach (var id in ids)
                {
                    //ObjectId id = ids[0];
                    //Polyline objPl = tr.GetObject(id, OpenMode.ForRead) as Polyline;
                    BlockReference objPrancha = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                    Extents3d bounds = (Extents3d)objPrancha.Bounds.Value;
                    Point2d ptMin = new Point2d(bounds.MinPoint.X, bounds.MinPoint.Y);
                    Point2d ptMax = new Point2d(bounds.MaxPoint.X, bounds.MaxPoint.Y);
                    Extents2d plotWin = new Extents2d(ptMin, ptMax);

                    GetSheetExtents(tr, objPrancha);

                    try
                    {
                        PlotSettingsValidator psv = PlotSettingsValidator.Current;
                        psv.SetPlotWindowArea(plotSettingsObj, plotWin);
                        psv.SetPlotRotation(plotSettingsObj, PlotRotation.Degrees180);

                        //now plot the setup...
                        PlotInfo plotInfo = new PlotInfo();
                        plotInfo.Layout = layoutObj.ObjectId;
                        plotInfo.OverrideSettings = plotSettingsObj;
                        PlotInfoValidator piv = new PlotInfoValidator();
                        piv.Validate(plotInfo);

                        string outName = "C:\\Users\\weiglas.ribeiro.LANGAMER\\Desktop\\" + System.DateTime.Now.Millisecond + ".pdf";

                        using (var pe = PlotFactory.CreatePublishEngine())
                        {
                            // Begin plotting a document.
                            pe.BeginPlot(null, null);
                            pe.BeginDocument(plotInfo, doc.Name, null, 1, true, outName);

                            // Begin plotting the page
                            PlotPageInfo ppi = new PlotPageInfo();
                            pe.BeginPage(ppi, plotInfo, true, null);
                            pe.BeginGenerateGraphics(null);
                            pe.EndGenerateGraphics(null);

                            // Finish the sheet
                            pe.EndPage(null);

                            // Finish the document
                            pe.EndDocument(null);

                            //// And finish the plot
                            pe.EndPlot(null);
                        }
                    }
                    catch(Exception er)
                    {

                    }
                }
                tr.Commit();
            }
        }

        private static Entity cachedEnt = null;
        private static ObjectId cachedId = ObjectId.Null;
        private static Entity OpenObject(Transaction tr, ObjectId id)
        {
            if (cachedId != id || cachedEnt == null)
                cachedEnt = tr.GetObject(id, OpenMode.ForRead) as Entity;
            return cachedEnt;
        }

        private static Extents2d GetSheetExtents(Transaction tr, BlockReference blkSheet)
        {
            List<ObjectId> obj;
            using (DBObjectCollection dbObjCol = new DBObjectCollection())
            {
                blkSheet.Explode(dbObjCol);
                IEnumerable<ObjectId> b = dbObjCol.Cast<ObjectId>();
                obj = (from obj1 in b
                           where obj1.ObjectClass.DxfName.ToString() == "LWPOLYLINE"
                           select obj1).ToList<ObjectId>();
            }

                return new Extents2d(new Point2d(0, 0), new Point2d(1, 1));
        }
    }
}


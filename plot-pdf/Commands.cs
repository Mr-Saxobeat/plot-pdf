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
                       where id1.ObjectClass.DxfName.ToUpper() == "LWPOLYLINE"
                       select id1).ToList<ObjectId>();
                // Fim: Pega todos os POLYLINE do desenho atual **********************************************************************************************

                //foreach (var id in ids)
                //{
                    ObjectId id = ids[0];
                    Polyline objPl = tr.GetObject(id, OpenMode.ForRead) as Polyline;
                    Point2d pMin = objPl.GetPoint2dAt(0);
                    Point2d pMax = objPl.GetPoint2dAt(2);
                    Extents2d plotWin= new Extents2d(pMin, pMax);

                    try
                    {
                        PlotSettingsValidator psv = PlotSettingsValidator.Current;
                        psv.SetPlotWindowArea(plotSettingsObj, plotWin);
                        
                        //now plot the setup...
                        PlotInfo plotInfo = new PlotInfo();
                        plotInfo.Layout = layoutObj.ObjectId;
                        plotInfo.OverrideSettings = plotSettingsObj;
                        PlotInfoValidator piv = new PlotInfoValidator();
                        piv.Validate(plotInfo);

                        string outName = "C:\\temp\\" + plotSettingsObj.PlotSettingsName;

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
                //}

                tr.Commit();
            }
        }



        [CommandMethod("PlotPageSetup")]
        static public void PlotPageSetup()

        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptStringOptions opts = new PromptStringOptions("Enter plot setting name");
            opts.AllowSpaces = true;
            PromptResult settingName = ed.GetString(opts);

            if (settingName.Status != PromptStatus.OK)
                return;

            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                LayoutManager layManager = LayoutManager.Current;
                ObjectId layoutId = layManager.GetLayoutId(layManager.CurrentLayout);
                Layout layoutObj = (Layout)Tx.GetObject(layoutId, OpenMode.ForRead);
                DBDictionary plotSettingsDic = (DBDictionary)Tx.GetObject(db.PlotSettingsDictionaryId, OpenMode.ForRead);
                

                if (!plotSettingsDic.Contains(settingName.StringResult))
                    return;

                ObjectId plotsetting = plotSettingsDic.GetAt(settingName.StringResult);

                //layout type
                bool bModel = layoutObj.ModelType;
                PlotSettings plotSettings = Tx.GetObject(plotsetting, OpenMode.ForRead) as PlotSettings;
    
                if (plotSettings.ModelType != bModel)
                    return;

                object backgroundPlot = Application.GetSystemVariable("BACKGROUNDPLOT");
                Application.SetSystemVariable("BACKGROUNDPLOT", 0);

                try
                {
                    //now plot the setup...
                    PlotInfo plotInfo = new PlotInfo();
                    plotInfo.Layout = layoutObj.ObjectId;
                    plotInfo.OverrideSettings = plotSettings;
                    PlotInfoValidator piv = new PlotInfoValidator();
                    piv.Validate(plotInfo);

                    string outName = "c:\\temp\\" + plotSettings.PlotSettingsName;

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
                catch
                {

                }
                Tx.Commit();

                Application.SetSystemVariable("BACKGROUNDPLOT", backgroundPlot);//
            }
        }
    }
}

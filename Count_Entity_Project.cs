using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;

[assembly: CommandClass(typeof(EntityCounter.CountCommand))]

namespace EntityCounter
{
    public class CountCommand
    {
        [CommandMethod("CountEntitiesMulti")]
        public void CountEntitiesMulti()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            ed.WriteMessage("\n--- MULTIPLE ENTITY COUNTER ---");

            // Select entities
            PromptSelectionResult selRes = ed.GetSelection();
            if (selRes.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nNo entities selected.");
                return;
            }

            // Table insertion point
            PromptPointResult ptRes = ed.GetPoint("\nSelect table insertion point: ");
            if (ptRes.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nCancelled.");
                return;
            }

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    Dictionary<string, int> entityCounts = new Dictionary<string, int>();

                    foreach (SelectedObject so in selRes.Value)
                    {
                        if (so == null) continue;

                        Entity ent = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Entity;
                        if (ent == null) continue;

                        string key = GetKey(ent, tr);

                        if (entityCounts.ContainsKey(key))
                            entityCounts[key]++;
                        else
                            entityCounts.Add(key, 1);
                    }

                    CreateTable(db, tr, entityCounts, ptRes.Value);

                    ed.WriteMessage("\nCounting completed successfully.");

                    tr.Commit();
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage("\nError: " + ex.Message);
                    tr.Abort();
                }
            }
        }

        // Get unique key for entity
        private string GetKey(Entity ent, Transaction tr)
        {
            if (ent is BlockReference br)
            {
                return "BLOCK_" + GetBlockName(br, tr).ToUpper();
            }

            return ent.GetType().Name.ToUpper();
        }

        // Get block name safely
        private string GetBlockName(BlockReference br, Transaction tr)
        {
            try
            {
                if (br.IsDynamicBlock)
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead);
                    return btr.Name;
                }
                else
                {
                    return br.Name;
                }
            }
            catch
            {
                return "UNKNOWN_BLOCK";
            }
        }

        // Create table
        private void CreateTable(Database db, Transaction tr,
            Dictionary<string, int> data, Point3d pt)
        {
            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            BlockTableRecord ms =
                (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

            Table tb = new Table();

            tb.TableStyle = db.Tablestyle;
            tb.Position = pt;

            int rows = data.Count + 2;
            int cols = 2;

            tb.SetSize(rows, cols);
            tb.SetRowHeight(10);
            tb.SetColumnWidth(60);

            // Title
            tb.Cells[0, 0].TextString = "ENTITY COUNT";
            tb.MergeCells(CellRange.Create(tb, 0, 0, 0, 1));
            tb.Cells[0, 0].Alignment = CellAlignment.MiddleCenter;

            // Headers
            tb.Cells[1, 0].TextString = "Entity";
            tb.Cells[1, 1].TextString = "Count";

            // Data
            int row = 2;
            foreach (KeyValuePair<string, int> item in data)
            {
                tb.Cells[row, 0].TextString = item.Key;
                tb.Cells[row, 1].TextString = item.Value.ToString();
                row++;
            }

            tb.GenerateLayout();

            ms.AppendEntity(tb);
            tr.AddNewlyCreatedDBObject(tb, true);
        }
    }
}
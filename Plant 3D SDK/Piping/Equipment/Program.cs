//
// (C) Copyright 2014 by Autodesk, Inc.
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is subject to
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable.
//


using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Collections;

// Platform
//
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Windows;

// Plant
//
using Autodesk.ProcessPower.PnP3dObjects;
using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.ProjectManager;
using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.DataObjects;
using Autodesk.ProcessPower.PnP3dDataLinks;
using Autodesk.ProcessPower.PartsRepository;
using Autodesk.ProcessPower.PnP3dEquipment;
using Autodesk.ProcessPower.ACPUtils;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using PlantApp = Autodesk.ProcessPower.PlantInstance.PlantApplication;

[assembly: Autodesk.AutoCAD.Runtime.ExtensionApplication(null)]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(EquipmentSample.Program))]

namespace EquipmentSample
{
    public class Program
    {
        // Currently selected equipment
        //
        static EquipmentType    s_CurrentEquipmentType = null;
        static String           s_CurrentDwgName = null;
        static double           s_CurrentDwgScale = 1.0;
        static ObjectId         s_CurrentEquipmentId = ObjectId.Null;


        // Load equipment package from content folder
        //
        [CommandMethod("EquipmentLoadPackage")]
        public static void EquipmentLoadPackage()
        {
            try
            {
                Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;

                // Use helper class
                //
                using (EquipmentHelper eqHelper = new EquipmentHelper())
                {
                    // All available packages
                    //
                    List<String> kwds = new List<String>();
                    int cnt = 0;
                    foreach (EquipmentType eqt in eqHelper.EquipmentPackages)
                    {
                        // Make a keyword in the form "n.DisplayName"
                        // And kill the spaces
                        //
                        cnt += 1;
                        kwds.Add(cnt.ToString() + "." + eqt.DisplayName.Replace(" ", ""));
                    }

                    if (kwds.Count == 0)
                    {
                        ed.WriteMessage("\nNo equipment packages found");
                        return;
                    }

                    // Ask
                    //
                    PromptResult res = ed.GetKeywords("\nSelect equipment type", kwds.ToArray());
                    if (res.Status == PromptStatus.OK)
                    {
                        // Find selected
                        //
                        int j = kwds.IndexOf(res.StringResult);
                        EquipmentType eqt = eqHelper.EquipmentPackages[j];

                        // Scale and init for the project, and make current
                        //
                        s_CurrentEquipmentType = eqHelper.MakeEquipmentForProject(eqt);
                        s_CurrentDwgName = null;
                        s_CurrentEquipmentId = ObjectId.Null;
                    }
                }
            }
            catch (System.Exception)
            {
            }
        }


        // Load equipment template
        //
        [CommandMethod("EquipmentLoadTemplate")]
        public static void EquipmentLoadTemplate()
        {
            try
            {
                Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;

                // Use helper class
                //
                using (EquipmentHelper eqHelper = new EquipmentHelper())
                {
                    // All available templates
                    //
                    StringCollection tmpls = eqHelper.GetEquipmentTemplates(eqHelper.EquipmentTemplateFolder);
                    List<String> kwds = new List<String>();
                    int cnt = 0;
                    foreach (String s in tmpls)
                    {
                        // Make a keyword in the form "n.Name"
                        // And kill the spaces
                        //
                        cnt += 1;
                        String fname = System.IO.Path.GetFileNameWithoutExtension(s);
                        kwds.Add(cnt.ToString() + "." + fname.Replace(" ", ""));
                    }

                    if (kwds.Count == 0)
                    {
                        ed.WriteMessage("\nNo equipment templates found");
                        return;
                    }

                    // Ask
                    //
                    PromptResult res = ed.GetKeywords("\nSelect equipment type", kwds.ToArray());
                    if (res.Status == PromptStatus.OK)
                    {
                        // Find selected and load
                        //
                        int j = kwds.IndexOf(res.StringResult);
                        String dwgName = null;
                        double dwgScale = 1.0;
                        EquipmentType eqt = eqHelper.LoadTemplate(tmpls[j], out dwgName, out dwgScale);

                        // Scale and init for the project, and make current
                        //
                        s_CurrentEquipmentType = eqHelper.MakeEquipmentForProject(eqt);
                        s_CurrentDwgName = dwgName;
                        s_CurrentDwgScale = dwgScale;
                        s_CurrentEquipmentId = ObjectId.Null;
                    }
                }
            }
            catch (System.Exception)
            {
            }
        }


        // Load equipment from entity
        //
        [CommandMethod("EquipmentLoadEntity")]
        public static void EquipmentLoadEntity()
        {
            try
            {
                Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;

                // Select entity
                //
                ObjectId eqId;
                while (true)
                {
                    PromptEntityResult res = ed.GetEntity("\nSelect equipment entity: ");
                    if (res.Status == PromptStatus.OK)
                    {
                        // Equipment ?
                        //
                        if (res.ObjectId.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Equipment))))
                        {
                            // Yes
                            //
                            eqId = res.ObjectId;
                            break;
                        }
                    }
                    else
                    if (res.Status == PromptStatus.Cancel)
                    {
                        return;
                    }
                }

                // Use helper class
                //
                using (EquipmentHelper eqHelper = new EquipmentHelper())
                {
                    // Load
                    //
                    s_CurrentEquipmentType = eqHelper.RetrieveEquipmentFromInstance(eqId);
                    s_CurrentDwgName = null;
                    s_CurrentEquipmentId = eqId;
                }
            }
            catch (System.Exception)
            {
            }
        }


        // Create equipment entity
        //
        [CommandMethod("EquipmentCreate")]
        public static void EquipmentCreate()
        {
            try
            {
                Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;
                Database db                 = AcadApp.DocumentManager.MdiActiveDocument.Database;
                Project currentProject      = PlantApp.CurrentProject.ProjectParts["Piping"];
                DataLinksManager dlm        = currentProject.DataLinksManager;
                DataLinksManager3d dlm3d    = DataLinksManager3d.Get3dManager(dlm);

                if (s_CurrentEquipmentType == null)
                {
                    ed.WriteMessage("\nNo current equipment type loaded");
                    return;
                }

                if (s_CurrentEquipmentType.IsImportedEquipment())
                {
                    // Imported equipment has to be a template
                    //
                    if (s_CurrentDwgName == null)
                    {
                        ed.WriteMessage("\nImported equipment can be only used as a template.");
                        return;
                    }
                }

                // Use helper class
                //
                using (EquipmentHelper eqHelper = new EquipmentHelper())
                {
                    // Clear tag values, just in case, to avoid duplicated tags, if used the second time
                    //
                    eqHelper.InitTagRelatedProperties(s_CurrentEquipmentType);

                    // Create equipment entity
                    //
                    PartSizeProperties equipPart;
                    PartSizePropertiesCollection nozzleParts;
                    Equipment eqEnt = eqHelper.CreateEquipmentEntity (s_CurrentEquipmentType,
                                                                      s_CurrentDwgName,
                                                                      s_CurrentDwgScale,
                                                                      db,
                                                                      out equipPart,
                                                                      out nozzleParts);
                    if (eqEnt == null)
                    {
                        ed.WriteMessage("\nCan't create equipment entity");
                        return;
                    }

                    // Drag. Ask for a placement point and rotation angle
                    //
                    EqDragPoint dragPoint = new EqDragPoint(eqEnt);
                    if (ed.Drag(dragPoint).Status != PromptStatus.OK)
                    {
                        eqEnt.Dispose();
                        return;
                    }
                    EqDragAngle dragAngle = new EqDragAngle(eqEnt);
                    ed.Drag(dragAngle);

                    // Add
                    //
                    using (Transaction tr = db.TransactionManager.StartOpenCloseTransaction())
                    {
                        using (PipingObjectAdder pipeObjAdder = new PipingObjectAdder(dlm3d, db))
                        {
                            pipeObjAdder.Add(equipPart, eqEnt, nozzleParts);
                            tr.AddNewlyCreatedDBObject(eqEnt, true);

                            // Store entity
                            //
                            s_CurrentEquipmentId = eqEnt.ObjectId;
                        }

                        tr.Commit();
                    }
                }
            }
            catch (System.Exception)
            {
            }
        }


        // Convert equipment
        //
        [CommandMethod("EquipmentConvert")]
        public static void EquipmentConvert()
        {
            try
            {
                Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;
                Database db                 = AcadApp.DocumentManager.MdiActiveDocument.Database;
                Project currentProject      = PlantApp.CurrentProject.ProjectParts["Piping"];
                DataLinksManager dlm        = currentProject.DataLinksManager;
                PnPDatabase pnpDb           = dlm.GetPnPDatabase();
                DataLinksManager3d dlm3d    = DataLinksManager3d.Get3dManager(dlm);


                // Existing equipment classes
                // TODO: need to recursively get derived from derived...
                //
                PnPTable eqTable = pnpDb.Tables["Equipment"];
                StringCollection tbls = eqTable.DerivedTables;

                // Form kwd list
                //
                List<String> kwds = new List<String>();
                int cnt = 0;
                foreach (String s in tbls)
                {
                    // Make a keyword in the form "n.DisplayName"
                    // And kill the spaces
                    //
                    cnt += 1;
                    kwds.Add(cnt.ToString() + "." + s.Replace(" ", ""));
                }

                // Ask to select a class
                //
                String eqClass = null;
                PromptResult res = ed.GetKeywords("\nSelect equipment class", kwds.ToArray());
                if (res.Status == PromptStatus.OK)
                {
                    // Find selected
                    //
                    int j = kwds.IndexOf(res.StringResult);
                    eqClass = tbls[j];
                }
                else
                {
                    return;
                }


                // Select ents to convert
                //
                PromptSelectionOptions opts = new PromptSelectionOptions();
                PromptSelectionResult selRes = ed.GetSelection(opts);
                if (selRes.Status != PromptStatus.OK)
                {
                    return;
                }
                SelectionSet ss = selRes.Value;
                if (ss.Count == 0)
                {
                    return;
                }


                // Ask for a base point
                //
                Point3d basePoint;
                PromptPointResult pRes = ed.GetPoint("\nSelect base point: ");
                if (pRes.Status != PromptStatus.OK)
                {
                    return;
                }
                basePoint = pRes.Value;


                // Use helper class
                //
                using (EquipmentHelper eqHelper = new EquipmentHelper())
                {
                    // New type
                    //
                    EquipmentType eqt = eqHelper.NewImportedProjectEquipment(eqClass);

                    // Convert
                    //
                    PartSizeProperties equipPart;
                    Equipment eqEnt = eqHelper.CreateEquipmentEntity (eqt,
                                                                      ss.GetObjectIds(),
                                                                      basePoint,
                                                                      out equipPart);
                    if (eqEnt == null)
                    {
                        ed.WriteMessage("\nCan't convert equipment entity");
                        return;
                    }

                    // Drag. Ask for a placement point and rotation angle
                    //
                    EqDragPoint dragPoint = new EqDragPoint(eqEnt);
                    if (ed.Drag(dragPoint).Status != PromptStatus.OK)
                    {
                        eqEnt.Dispose();
                        return;
                    }
                    EqDragAngle dragAngle = new EqDragAngle(eqEnt);
                    ed.Drag(dragAngle);

                    // Add
                    //
                    using (Transaction tr = db.TransactionManager.StartOpenCloseTransaction())
                    {
                        using (PipingObjectAdder pipeObjAdder = new PipingObjectAdder(dlm3d, db))
                        {
                            pipeObjAdder.Add(equipPart, eqEnt, null);
                            tr.AddNewlyCreatedDBObject(eqEnt, true);

                            // Store entity and equipment
                            //
                            s_CurrentEquipmentType = eqt;
                            s_CurrentDwgName = null;
                            s_CurrentEquipmentId = eqEnt.ObjectId;
                        }

                        tr.Commit();
                    }
                }
            }
            catch (System.Exception)
            {
            }
        }


        // Modify equipment (and enitity if selected)
        //
        [CommandMethod("EquipmentModify")]
        public static void EquipmentModify()
        {
            try
            {
                Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;

                if (s_CurrentEquipmentType == null)
                {
                    ed.WriteMessage("\nNo current equipment type loaded");
                    return;
                }

                if (s_CurrentEquipmentType.IsImportedEquipment())
                {
                    // Need an entity to edit
                    //
                    if (s_CurrentEquipmentId.IsNull)
                    {
                        ed.WriteMessage("\nSelect equipment entity to edit imported equipment.");
                        return;
                    }
                }

                // Use helper class
                //
                using (EquipmentHelper eqHelper = new EquipmentHelper())
                {
                    // Form kwords
                    //
                    PromptKeywordOptions opts = new PromptKeywordOptions("\nSelect option:");
                    if (!s_CurrentEquipmentId.IsNull)
                    {
                        // Attach/Detach graphics to the current entity
                        //
                        opts.Keywords.Add("Attach");
                        opts.Keywords.Add("Detach");
                    }
                    if (s_CurrentEquipmentType.IsFabricatedEquipment())
                    {
                        // Edit geometrical shapes
                        //
                        opts.Keywords.Add("Shapes");
                    }
                    else
                    if (s_CurrentEquipmentType.IsParametricEquipment())
                    {
                        // Edit parameters
                        //
                        opts.Keywords.Add("Parameters");
                    }
                    opts.Keywords.Add("eXit");

                    // Ask
                    //
                    while (true)
                    {
                        PromptResult res = ed.GetKeywords(opts);
                        if (res.Status != PromptStatus.OK)
                        {
                            return;
                        }
                        if (res.StringResult == "Attach")
                        {
                            AttachGraphics(eqHelper);
                        }
                        else
                        if (res.StringResult == "Detach")
                        {
                            DetachGraphics(eqHelper);
                        }
                        else
                        if (res.StringResult == "Shapes")
                        {
                            EditShapes(eqHelper);
                        }
                        else
                        if (res.StringResult == "Parameters")
                        {
                            // Collect all the parameters from the current equipment
                            //
                            ParameterList plist = new ParameterList();
                            plist.AddRange(s_CurrentEquipmentType.Parameters);

                            // And geom categories params
                            //
                            foreach (CategoryInfo ci in s_CurrentEquipmentType.Categories)
                            {
                                plist.AddRange(ci.Parameters);
                            }

                            // And nozzle params
                            //
                            foreach (NozzleInfo ni in s_CurrentEquipmentType.Nozzles)
                            {
                                plist.AddRange(ni.Parameters);
                            }

                            // Edit now
                            //
                            EditParameters(eqHelper, plist);
                        }
                        else
                        {
                            return;
                        }
                    }
                }
            }
            catch (System.Exception)
            {
            }
        }

        static void AttachGraphics (EquipmentHelper eqHelper)
        {
            Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;

            // Select ents to convert
            //
            PromptSelectionOptions opts = new PromptSelectionOptions();
            PromptSelectionResult selRes = ed.GetSelection(opts);
            if (selRes.Status != PromptStatus.OK)
            {
                return;
            }
            SelectionSet ss = selRes.Value;
            if (ss.Count == 0)
            {
                return;
            }

            // Attach
            //
            eqHelper.AttachGraphics(s_CurrentEquipmentId, ss.GetObjectIds());
        }

        static void DetachGraphics (EquipmentHelper eqHelper)
        {
            // Anything attached?
            //
            ObjectId[] ids = eqHelper.FindAttachedGraphics(s_CurrentEquipmentId);
            if (ids.Length == 0)
            {
                Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage("\nCurrent equipment has no attached graphics.");
            }
            else
            {
                eqHelper.DetachGraphics(s_CurrentEquipmentId, ids);
            }
        }

        static void EditShapes (EquipmentHelper eqHelper)
        {
            Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;
            CategoryInfo curCategory    = null;

            while (true)
            {
                // Options
                //
                PromptKeywordOptions opts = new PromptKeywordOptions("\nSelect Option:");
                if (s_CurrentEquipmentType.Categories.Count > 0)
                {
                    opts.Keywords.Add("Select");
                }
                opts.Keywords.Add("Add");
                if (curCategory != null)
                {
                    opts.Keywords.Add("Remove");
                    opts.Keywords.Add("Parameters");
                }
                opts.Keywords.Add("eXit");

                // Ask
                //
                PromptResult res = ed.GetKeywords(opts);
                if (res.Status != PromptStatus.OK)
                {
                    return;
                }
                if (res.StringResult == "Select")
                {
                    // Select current
                    //
                    PromptKeywordOptions opts1 = new PromptKeywordOptions(/*NOXLATE*/"\nSelect Shape:");
                    for (int j = 0; j < s_CurrentEquipmentType.Categories.Count; j++)
                    {
                        CategoryInfo ci = s_CurrentEquipmentType.Categories[j];
                        opts1.Keywords.Add(String.Format("{0}.{1}", (j + 1).ToString(), ci.ToString().Replace(/*NOXLATE*/" ",/*NOXLATE*/"")));
                    }
                    PromptResult res1 = ed.GetKeywords(opts1);
                    if (res1.Status == PromptStatus.OK)
                    {
                        curCategory = null;
                        for (int j = 0; j < opts1.Keywords.Count; j++)
                        {
                            if (res1.StringResult == opts1.Keywords[j].GlobalName)
                            {
                                curCategory = s_CurrentEquipmentType.Categories[j];
                                break;
                            }
                        }
                    }
                }
                else
                if (res.StringResult == "Add")
                {
                    // Add new shape
                    //
                    PromptKeywordOptions opts1 = new PromptKeywordOptions(/*NOXLATE*/"\nSelect Shape to add:");
                    EquipmentPrimitives primitives = EquipmentPrimitives.GetPrimitives();
                    List<String> keylist = new List<String>();
                    foreach (String key in primitives.Keys)
                    {
                        String dname = primitives[key].ToString().Replace(" ", "").ToUpper();
                        keylist.Add(key);
                        opts1.Keywords.Add(key, dname);
                    }
                    PromptResult res1 = ed.GetKeywords(opts1);
                    if (res1.Status == PromptStatus.OK)
                    {
                        for (int j = 0; j < opts1.Keywords.Count; j++)
                        {
                            if (res1.StringResult == opts1.Keywords[j].GlobalName)
                            {
                                // Add
                                //
                                CategoryInfo ci = primitives[keylist[j]];
                                s_CurrentEquipmentType.Categories.Add(ci.Clone() as CategoryInfo);

                                // Update entity
                                //
                                if (!s_CurrentEquipmentId.IsNull)
                                {
                                    eqHelper.UpdateEquipmentEntity(s_CurrentEquipmentId, s_CurrentEquipmentType, s_CurrentDwgName, s_CurrentDwgScale);
                                }
                                break;
                            }
                        }
                    }
                }
                else
                if (res.StringResult == "Remove")
                {
                    // Remove from equipment
                    //
                    s_CurrentEquipmentType.Categories.Remove(curCategory);
                    curCategory = null;

                    // Update entity
                    //
                    if (!s_CurrentEquipmentId.IsNull)
                    {
                        eqHelper.UpdateEquipmentEntity(s_CurrentEquipmentId, s_CurrentEquipmentType, s_CurrentDwgName, s_CurrentDwgScale);
                    }
                }
                else
                if (res.StringResult == "Parameters")
                {
                    EditParameters(eqHelper, curCategory.Parameters);
                }
                else
                {
                    return;
                }
            }
        }

        static void EditParameters (EquipmentHelper eqHelper, ParameterList plist)
        {
            Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;

            // Ask for a parameter to edit
            //
            PromptKeywordOptions opts = new PromptKeywordOptions("\nSelect Parameter:");
            while(true)
            {
                // Kwords
                //
                opts.Keywords.Clear();
                foreach (ParameterInfo pi in plist)
                {
                    // Kill underscore
                    //
                    String locKwd = pi.Name.Replace("_", "");
                    String globKwd = pi.ToDefinitionString().Replace("_", "");
                    opts.Keywords.Add(locKwd, globKwd);
                }
                opts.Keywords.Add("eXit");

                // Ask
                //
                PromptResult res = ed.GetKeywords(opts);
                if (res.Status != PromptStatus.OK)
                {
                    return;
                }
                if (res.StringResult == /*NOXLATE*/"eXit")
                {
                    return;
                }

                // Find selected parameter
                //
                for (int j = 0; j < opts.Keywords.Count; j++)
                {
                    if (res.StringResult == opts.Keywords[j].GlobalName)
                    {
                        ParameterInfo pi = plist[j];

                        // Ask for a new value
                        //
                        PromptStringOptions sopts = new PromptStringOptions(String.Format("{0}=", pi.Name));
                        sopts.DefaultValue = pi.Value;
                        PromptResult sres = ed.GetString(sopts);

                        if (sres.Status == PromptStatus.OK)
                        {
                            pi.Value = sres.StringResult;

                            // Todo: validate
                            //

                            // Update entity
                            //
                            if (!s_CurrentEquipmentId.IsNull)
                            {
                                eqHelper.UpdateEquipmentEntity(s_CurrentEquipmentId, s_CurrentEquipmentType, s_CurrentDwgName, s_CurrentDwgScale);
                            }
                        }
                        break;
                    }
                }
            }
        }


        // Save equipment template
        //
        [CommandMethod("EquipmentSaveTemplate")]
        public static void EquipmentSaveTemplate()
        {
            try
            {
                Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;
                Database db                 = AcadApp.DocumentManager.MdiActiveDocument.Database;

                if (s_CurrentEquipmentType == null)
                {
                    ed.WriteMessage("\nNo current equipment type loaded");
                    return;
                }

                // Use helper class
                //
                using (EquipmentHelper eqHelper = new EquipmentHelper())
                {
                    // We nay need a block with the graphics
                    //
                    ObjectId blockId = ObjectId.Null;
                    
                    if (!s_CurrentEquipmentId.IsNull)
                    {
                        // Get it from entity
                        //
                        try
                        {
                            using (Transaction tr = db.TransactionManager.StartOpenCloseTransaction())
                            {
                                Equipment eqEnt = tr.GetObject(s_CurrentEquipmentId, OpenMode.ForRead) as Equipment;
                                if (eqEnt != null)
                                {
                                    blockId = eqEnt.SymbolId;
                                }

                                tr.Commit();
                            }
                        }
                        catch
                        {
                        }
                    }

                    if (blockId.IsNull && s_CurrentDwgName != null)
                    {
                        // Use dwg file with graphics
                        //
                        try
                        {
                            // Create temporary equipment entity and find its block
                            //
                            PartSizeProperties equipPart;
                            PartSizePropertiesCollection nozzleParts;
                            Equipment eqEnt = eqHelper.CreateEquipmentEntity (s_CurrentEquipmentType,
                                                                              s_CurrentDwgName,
                                                                              s_CurrentDwgScale,
                                                                              db,
                                                                              out equipPart,
                                                                              out nozzleParts);
                            if (eqEnt != null)
                            {
                                blockId = eqEnt.SymbolId;
                                eqEnt.Dispose();
                            }
                        }
                        catch
                        {
                        }
                    }

                    if (blockId.IsNull)
                    {
                        // Can't save imported equipment without a block. It would have no graphics
                        //
                        if (s_CurrentEquipmentType.IsImportedEquipment())
                        {
                            ed.WriteMessage("\nImported equipment needs an entity to be saved as a template.");
                            return;
                        }
                    }

                    // Ask for a file name
                    //
    				String cwd = System.IO.Directory.GetCurrentDirectory();

                    String defaultName = System.IO.Path.Combine(eqHelper.EquipmentTemplateFolder, s_CurrentEquipmentType.DisplayName);
				    SaveFileDialog dlg = new SaveFileDialog("Save Template As",
    					                                    defaultName,
    					                                    eqHelper.EquipmentTemplateExtension,
                                                            "SaveTemplate",
                                                    		SaveFileDialog.SaveFileDialogFlags.DoNotTransferRemoteFiles
                                                    		| SaveFileDialog.SaveFileDialogFlags.ForceDefaultFolder
                                                    		| SaveFileDialog.SaveFileDialogFlags.NoFtpSites
                                                    		| SaveFileDialog.SaveFileDialogFlags.NoUrls);
                    System.Windows.Forms.DialogResult res = dlg.ShowDialog();

				    System.IO.Directory.SetCurrentDirectory(cwd);

    				if (res == System.Windows.Forms.DialogResult.OK)
                    {
                        eqHelper.SaveTemplate(s_CurrentEquipmentType, dlg.Filename, blockId);
                    }
                }
            }
            catch (System.Exception)
            {
            }
        }


        // Add nozzle
        //
        [CommandMethod("EquipmentAddNozzle")]
        public static void EquipmentAddNozzle()
        {
            try
            {
                Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;
                Database db                 = AcadApp.DocumentManager.MdiActiveDocument.Database;

                if (s_CurrentEquipmentType == null)
                {
                    ed.WriteMessage("\nNo current equipment type loaded");
                    return;
                }

                bool bFabricated = false;
                if (s_CurrentEquipmentType.IsImportedEquipment())
                {
                    // Need an entity to edit
                    //
                    if (s_CurrentEquipmentId.IsNull)
                    {
                        ed.WriteMessage("\nSelect equipment entity to edit imported equipment.");
                        return;
                    }
                }
                else
                if (s_CurrentEquipmentType.IsParametricEquipment())
                {
                    // Can't add nozzle to parametric
                    //
                    ed.WriteMessage("\nCan't add nozzle to parametic equipment.");
                    return;
                }
                else
                {
                    bFabricated = true;
                }

                // Use helper class
                //
                using (EquipmentHelper eqHelper = new EquipmentHelper())
                {
                    // New Nozzle
                    //
                    int newIndex = eqHelper.NewNozzleIndex(s_CurrentEquipmentType);
                    String nozName = "New Nozzle " + newIndex.ToString();
                    NozzleInfo ni = eqHelper.NewNozzle(s_CurrentEquipmentType, newIndex, nozName);

                    // Project units
                    //
                    Units nUnit = (Units)Autodesk.ProcessPower.AcPp3dObjectsUtils.ProjectUnits.CurrentNDUnit;
                    Units lUnit = (Units)Autodesk.ProcessPower.AcPp3dObjectsUtils.ProjectUnits.CurrentLinearUnit;

                    // Find nozzles
                    //
                    NominalDiameter nd = new NominalDiameter();
                    String pressureClass;
                    String facing;
                    nd.EUnits = nUnit;
                    if (nUnit == Units.Inch)
                    {
                        // Something hardcoded imperial
                        //
                        nd.Value = 4;
                        pressureClass = "300";
                        facing = "RF";
                    }
                    else
                    {
                        // Metric
                        //
                        nd.Value = 100;
                        pressureClass = "10";
                        facing = "C";
                    }

                    // For example, straight, flanged
                    //
                    PnPRow[] rows = NozzleInfo.SelectFromNozzleCatalog(eqHelper.NozzleRepository, "StraightNozzle", nd, "FL", pressureClass, facing);
                    if (rows.Length == 0)
                    {
                        ed.WriteMessage("\nNo nozzles found in the catalog.");
                        return;
                    }

                    // Take the first
                    // Its guid
                    //
                    String guid = String.Empty;
                    guid = rows[0][PartsRepository.PartGuid].ToString();

                    // Assign nozzle part
                    //
                    eqHelper.SetNozzlePart(s_CurrentEquipmentType, ni, guid);

                    if (bFabricated)
                    {
                        // TODO: we may ask for all these params

                        // For example, radial
                        //
                        eqHelper.SetNozzleLocation(s_CurrentEquipmentType, ni, (int)NozzleLocation.eRadial);

                        // Set some nozzle length: 1' / 300mm
                        //
                        ParameterInfo pa = ni.Parameters["L"];
                        if (pa != null)
                        {
                            if (lUnit == Units.Inch)
                            {
                                pa.Value = "12";
                            }
                            else
                            {
                                pa.Value = "300";
                            }
                        }

                        // And some height: 6"/150mm
                        //
                        pa = ni.Parameters["H"];
                        if (pa != null)
                        {
                            if (lUnit == Units.Inch)
                            {
                                pa.Value = "6";
                            }
                            else
                            {
                                pa.Value = "150";
                            }
                        }
                    }
                    else
                    {
                        // For imported equipment ask for a port position a direction
                        //
                        PromptPointResult res = ed.GetPoint("\nSelect port position: ");
                        if (res.Status != PromptStatus.OK)
                        {
                            return;
                        }
                        Point3d p = res.Value;
                        PromptPointOptions opts = new PromptPointOptions("\nSelect port direction: ");
                        opts.UseBasePoint = true;
                        opts.BasePoint = p;
                        opts.UseDashedLine = true;
                        while(true)
                        {
                            res = ed.GetPoint(opts);
                            if (res.Status != PromptStatus.OK)
                            {
                                return;
                            }
                            if (res.Value != p)
                            {
                                break;
                            }
                        }

                        // We need equipment ECS
                        //
                        Matrix3d ecs = Matrix3d.Identity;
                        try
                        {
                            using (Transaction tr = db.TransactionManager.StartOpenCloseTransaction())
                            {
                                Equipment eqEnt = tr.GetObject(s_CurrentEquipmentId, OpenMode.ForRead) as Equipment;
                                if (eqEnt != null)
                                {
                                    ecs = eqEnt.Ecs;
                                }

                                tr.Commit();
                            }
                        }
                        catch
                        {
                        }

                        // Set port
                        //
                        eqHelper.SetNozzleLocation(s_CurrentEquipmentType, ni, p, res.Value, 0.0, ecs);
                    }

                    // Add new nozzle
                    //
                    s_CurrentEquipmentType.Nozzles.Add(ni);

                    // Update entity
                    //
                    if (!s_CurrentEquipmentId.IsNull)
                    {
                        eqHelper.UpdateEquipmentEntity(s_CurrentEquipmentId, s_CurrentEquipmentType, s_CurrentDwgName, s_CurrentDwgScale);
                    }
                }
            }
            catch (System.Exception)
            {
            }
        }


        // Dragger classes
        //
        public class EqDragPoint : EntityJig
        {
            Point3d m_CurrentPos;
            JigPromptPointOptions m_Opts;

            public EqDragPoint (Equipment eqEnt) : base(eqEnt)
            {
                m_CurrentPos = eqEnt.Position;

                m_Opts = new JigPromptPointOptions("\nSelect point");
                m_Opts.UserInputControls = (UserInputControls.Accept3dCoordinates | UserInputControls.NoZeroResponseAccepted | UserInputControls.NoNegativeResponseAccepted);
            }

            protected override SamplerStatus Sampler(JigPrompts prompts)
            {
                // Ask for a point
                //
                PromptPointResult res = prompts.AcquirePoint(m_Opts);
                switch (res.Status)
                {
                case PromptStatus.OK:
                    if (m_CurrentPos == res.Value)
                    {
                        // The same point. Don't redraw
                        //
                        return SamplerStatus.NoChange;
                    }
                    else
                    {
                        // Take new point
                        //
                        m_CurrentPos = res.Value;
                        return SamplerStatus.OK;
                    }

                default:
                    return SamplerStatus.Cancel;
                }
            }

            protected override bool Update()
            {
                // Set new point
                //
                ((Equipment)Entity).Position = m_CurrentPos;
                return true;
            }
        }

        public class EqDragAngle : EntityJig
        {
            Double m_CurrentAngle;
            Double m_DefaultAngle;
            Point3d m_Position;
            Vector3d m_ZAxis;
            Matrix3d m_Trans;
            JigPromptAngleOptions m_Opts;

            public EqDragAngle (Equipment eqEnt) : base(eqEnt)
            {
                m_CurrentAngle = m_DefaultAngle = eqEnt.Rotation;
                m_Position = eqEnt.Position;
                m_ZAxis = eqEnt.ZAxis;
                m_Trans = Matrix3d.Identity;

                m_Opts = new JigPromptAngleOptions("\nSelect rotation angle");
                m_Opts.UserInputControls = UserInputControls.NullResponseAccepted;
                m_Opts.Cursor |= CursorType.RubberBand;
                m_Opts.BasePoint = eqEnt.Position;
                m_Opts.UseBasePoint = true;
                m_Opts.DefaultValue = m_DefaultAngle;
            }

            protected override SamplerStatus Sampler(JigPrompts prompts)
            {
                // Ask for an angle
                //
                PromptDoubleResult res = prompts.AcquireAngle(m_Opts);
                switch (res.Status)
                {
                case PromptStatus.OK:
                    if (m_CurrentAngle == res.Value)
                    {
                        // The same angle. Don't redraw
                        //
                        return SamplerStatus.NoChange;
                    }
                    else
                    {
                        // Rotate around equipment Z on the angle difference
                        //
                        m_Trans = Matrix3d.Rotation(res.Value-m_CurrentAngle, m_ZAxis, m_Position);
                        m_CurrentAngle = res.Value;
                    }
                    break;

                default:
                    // Restore default
                    //
                    m_Trans = Matrix3d.Rotation(m_DefaultAngle-m_CurrentAngle, m_ZAxis, m_Position);
                    m_CurrentAngle = m_DefaultAngle;
                    break;
                }

                // Transform here since Update isn't called on None or Cancel
                // And return NoChange to avoid second transform
                //
                Entity.TransformBy(m_Trans);
                return SamplerStatus.NoChange;
            }

            protected override bool Update()
            {
                // We transform in sampler()
                //
                return true;
            }
        }
    }
}

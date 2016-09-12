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
using Autodesk.ProcessPower.PnP3dPipeSupport;
using Autodesk.ProcessPower.P3dUI;
using Autodesk.ProcessPower.ACPUtils;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using PlantApp = Autodesk.ProcessPower.PlantInstance.PlantApplication;

[assembly: Autodesk.AutoCAD.Runtime.ExtensionApplication(null)]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(PipeSupportSample.Program))]

namespace PipeSupportSample
{
    public class Program
    {
        // Create support from script
        //
        [CommandMethod("SupportScriptCreate")]
        public static void SupportScriptCreate()
        {
            try
            {
                Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;
                Database db                 = AcadApp.DocumentManager.MdiActiveDocument.Database;

                // Ask for a script
                //
                List<SupportInfo> scripts = SupportHelper.Scripts;
                List<String> kwds = new List<String>();
                int cnt = 0;
                foreach (SupportInfo sup in scripts)
                {
                    // Make a keyword in the form "n.DisplayName"
                    // And kill spaces and other special symbols
                    //
                    cnt += 1;
                    String s = sup.ScriptDescription.Replace(" ", "");
                    s = s.Replace("/", "");
                    s = s.Replace(".", "");
                    kwds.Add(cnt.ToString() + "." + s);
                }

                if (kwds.Count == 0)
                {
                    ed.WriteMessage("\nNo scripts found");
                    return;
                }

                // Ask
                //
                PromptResult res = ed.GetKeywords("\nSelect support script", kwds.ToArray());
                if (res.Status != PromptStatus.OK)
                {
                    return;
                }

                // Find selected
                //
                int j = kwds.IndexOf(res.StringResult);
                SupportInfo info = scripts[j];

                // Use current size
                //
                UISettings sett = new UISettings();
                NominalDiameter nd = NominalDiameter.FromDisplayString(null, sett.CurrentSize);
                info.InitParameterList(nd);

                // Create entity
                //
                PartSizeProperties part = null;
                using (Support support = SupportHelper.CreateSupportEntity(info, db, out part))
                {
                    // Place
                    //
                    PlaceSupport(support, part);
                }
            }
            catch (System.Exception)
            {
            }
        }


        // Create support from support spec
        //
        [CommandMethod("SupportSpecCreate")]
        public static void SupportSpecCreate()
        {
            try
            {
                Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;
                Database db                 = AcadApp.DocumentManager.MdiActiveDocument.Database;

                // Use current size
                //
                UISettings sett = new UISettings();
                NominalDiameter nd = NominalDiameter.FromDisplayString(null, sett.CurrentSize);

                // Select from support spec
                //
                List<PartSizeProperties> parts = new List<PartSizeProperties>();
                SpecManager specMgr = SpecManager.GetSpecManager();
                SpecPartReader reader = specMgr.SelectParts("PipeSupportsSpec", "Support", nd);
                while (reader.Next())
                {
                    SpecPart part = new SpecPart();
                    part.AssignFrom(reader.Current);
                    parts.Add(part);
                }

                // Ask for a part
                //
                List<String> kwds = new List<String>();
                int cnt = 0;
                foreach (PartSizeProperties part in parts)
                {
                    // Make a keyword in the form "n.PartFamilyLongDescrtiption"
                    // And kill spaces and other special symbols
                    //
                    cnt += 1;
                    String s = "";
                    try
                    {
                        Object val = part.PropValue("PartFamilyLongDesc");
                        if (val != null)
                        {
                            s = val.ToString();
                            s = s.Replace(" ", "");
                            s = s.Replace("/", "");
                            s = s.Replace(".", "");
                        }
                    }
                    catch (System.Exception)
                    {
                    }
                    kwds.Add(cnt.ToString() + "." + s);
                }

                if (kwds.Count == 0)
                {
                    ed.WriteMessage("\nNo parts of the current size found");
                    return;
                }

                // Ask
                //
                PromptResult res = ed.GetKeywords("\nSelect support", kwds.ToArray());
                if (res.Status != PromptStatus.OK)
                {
                    return;
                }

                // Find selected
                //
                int j = kwds.IndexOf(res.StringResult);
                PartSizeProperties supportPart = parts[j];

                // Create entity
                //
                using (Support support = SupportHelper.CreateSupportEntity(supportPart, db))
                {
                    // Place
                    //
                    PlaceSupport(support, supportPart);
                }
            }
            catch (System.Exception)
            {
            }
        }


        // Create support from block
        //
        [CommandMethod("SupportBlockCreate")]
        public static void SupportBlockCreate()
        {
            try
            {
                Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;
                Database db                 = AcadApp.DocumentManager.MdiActiveDocument.Database;

                // Select block
                // It supposed to have ports already, added for example by PlantPartCOnvert command
                // If not, we would still create support, but without ports
                //
                String blockName = null;
                while (true)
                {
                    PromptEntityResult res = ed.GetEntity("\nSelect block: ");
                    if (res.Status == PromptStatus.OK)
                    {
                        // BlkRef ?
                        //
                        if (res.ObjectId.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(BlockReference))))
                        {
                            // Yes. Get block name
                            //
                            using (Transaction tr = db.TransactionManager.StartOpenCloseTransaction())
                            {
                                BlockReference blkRef = tr.GetObject(res.ObjectId, OpenMode.ForRead) as BlockReference;
                                blockName = blkRef.Name;
                                tr.Commit();
                            }
                            break;
                        }
                    }
                    else
                    if (res.Status == PromptStatus.Cancel)
                    {
                        return;
                    }
                }

                // Create custom script and init it with the current size
                //
                SupportInfo info = new SupportInfo();
                info.Type = SupportType.eCustom;
                UISettings sett = new UISettings();
                NominalDiameter nd = NominalDiameter.FromDisplayString(null, sett.CurrentSize);
                info.InitParameterList(nd);

                // Create support entity (null for dwgPath means current drawing)
                //
                PartSizeProperties part = null;
                using (Support support = SupportHelper.CreateSupportEntity(info, db, blockName, null, out part))
                {
                    // Place
                    //
                    PlaceSupport(support, part);
                }
            }
            catch (System.Exception)
            {
            }
        }


        // Create support from entities
        //
        [CommandMethod("SupportSelectionCreate")]
        public static void SupportSelectionCreate()
        {
            try
            {
                Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;
                Database db                 = AcadApp.DocumentManager.MdiActiveDocument.Database;

                // Select ents
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

                // Ask for a port position and direction
                //
                PromptPointResult res = ed.GetPoint("\nSelect port position: ");
                if (res.Status != PromptStatus.OK)
                {
                    return;
                }
                Point3d p = res.Value;
                PromptPointOptions pointOpts = new PromptPointOptions("\nSelect port direction: ");
                pointOpts.UseBasePoint = true;
                pointOpts.BasePoint = p;
                pointOpts.UseDashedLine = true;
                while(true)
                {
                    res = ed.GetPoint(pointOpts);
                    if (res.Status != PromptStatus.OK)
                    {
                        return;
                    }
                    if (res.Value != p)
                    {
                        break;
                    }
                }
                Vector3d dir = (res.Value - p).GetNormal();

                // Create custom script and init it with the current size
                //
                SupportInfo info = new SupportInfo();
                info.Type = SupportType.eCustom;
                UISettings sett = new UISettings();
                NominalDiameter nd = NominalDiameter.FromDisplayString(null, sett.CurrentSize);
                info.InitParameterList(nd);

                // Create support entity
                //
                PartSizeProperties part = null;
                using (Support support = SupportHelper.CreateSupportEntity(info, ss.GetObjectIds(), p, dir, out part))
                {
                    // Place
                    //
                    PlaceSupport(support, part);
                }
            }
            catch (System.Exception)
            {
            }
        }


        // Create support from entities
        //
        [CommandMethod("SupportConvert")]
        public static void SupportConvert()
        {
            try
            {
                Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;
                Database db                 = AcadApp.DocumentManager.MdiActiveDocument.Database;
                Project currentProject      = PlantApp.CurrentProject.ProjectParts["Piping"];
                DataLinksManager dlm        = currentProject.DataLinksManager;
                DataLinksManager3d dlm3d    = DataLinksManager3d.Get3dManager(dlm);

                // Select ents
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

                // Create custom script, init it with the current size, and make part
                //
                SupportInfo info = new SupportInfo();
                info.Type = SupportType.eCustom;
                UISettings sett = new UISettings();
                NominalDiameter nd = NominalDiameter.FromDisplayString(null, sett.CurrentSize);
                info.InitParameterList(nd);
                PartSizeProperties part = info.MakePart();

                // Select pipe or asset
                //
                ObjectId entId = ObjectId.Null;
                Autodesk.ProcessPower.PnP3dObjects.Port port = null;
                while (true)
                {
                    PromptEntityResult res = ed.GetEntity("\nSelect pipe or inline asset: ");
                    if (res.Status == PromptStatus.OK)
                    {
                        // Right type ?
                        //
                        if (res.ObjectId.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Pipe))) ||
                            res.ObjectId.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(PipeInlineAsset))))
                        {
                            // Yes. Can we snap to it ?
                            //
                            Point3d snapPoint = ed.Snap("Near", res.PickedPoint); 
                            if (SupportHelper.FindSupportAlignment(part, res.ObjectId, snapPoint, out port))
                            {
                                // Select it
                                //
                                entId = res.ObjectId;
                                break;
                            }
                        }
                    }
                    else
                    if (res.Status == PromptStatus.Cancel)
                    {
                        return;
                    }
                }

                // Resize with new nd
                // 
                info.InitParameterList(port.NominalDiameter);

                // Create support entity
                //
                Support support = SupportHelper.CreateSupportEntity(info, ss.GetObjectIds(), port.Position, -port.Direction, out part);
                if (support != null)
                {
                    // Align
                    //
                    if (!SupportHelper.AlignSupport(support, part, false, entId, port))
                    {
                        support.Dispose();
                        return;
                    }

                    // Add
                    //
                    ObjectId supId;
                    using (Transaction tr = db.TransactionManager.StartOpenCloseTransaction())
                    {
                        using (PipingObjectAdder pipeObjAdder = new PipingObjectAdder(dlm3d, db))
                        {
                            pipeObjAdder.Add(part, support);
                            tr.AddNewlyCreatedDBObject(support, true);
                            supId = support.ObjectId;
                        }

                        tr.Commit();
                    }

                    // Connect
                    //
                    SupportHelper.ConnectSupport(supId, entId);

                    // Cut if needed
                    //
                    SupportHelper.CutSupport(supId, entId);
                }
            }
            catch (System.Exception)
            {
            }
        }


        // Attach entities to support
        //
        [CommandMethod("SupportAttach")]
        public static void SupportAttach()
        {
            try
            {
                Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;

                // Select support
                //
                ObjectId id = ObjectId.Null;
                while (true)
                {
                    PromptEntityResult res = ed.GetEntity("\nSelect support: ");
                    if (res.Status == PromptStatus.OK)
                    {
                        // Support ?
                        //
                        if (res.ObjectId.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Support))))
                        {
                            // Yes. Take it
                            //
                            id = res.ObjectId;
                            break;
                        }
                    }
                    else
                    if (res.Status == PromptStatus.Cancel)
                    {
                        return;
                    }
                }

                // Select ents
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
                SupportHelper.AttachGraphics(id, ss.GetObjectIds());
            }
            catch (System.Exception)
            {
            }
        }


        // Detach
        //
        [CommandMethod("SupportDetach")]
        public static void SupportDetach()
        {
            try
            {
                Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;

                // Select support
                //
                ObjectId id = ObjectId.Null;
                while (true)
                {
                    PromptEntityResult res = ed.GetEntity("\nSelect support: ");
                    if (res.Status == PromptStatus.OK)
                    {
                        // Support ?
                        //
                        if (res.ObjectId.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Support))))
                        {
                            // Yes. Take it
                            //
                            id = res.ObjectId;
                            break;
                        }
                    }
                    else
                    if (res.Status == PromptStatus.Cancel)
                    {
                        return;
                    }
                }

                // Anything attached?
                //
                ObjectId[] ids = SupportHelper.FindAttachedGraphics(id);
                if (ids == null || ids.Length == 0)
                {
                    ed.WriteMessage("\nSupport has no attached graphics.");
                }
                else
                {
                    // Detach
                    //
                    SupportHelper.DetachGraphics(id, ids);
                }
            }
            catch (System.Exception)
            {
            }
        }


        // Edit
        //
        [CommandMethod("SupportEdit")]
        public static void SupportEdit()
        {
            try
            {
                Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;

                // Select support
                //
                ObjectId id = ObjectId.Null;
                while (true)
                {
                    PromptEntityResult res = ed.GetEntity("\nSelect support: ");
                    if (res.Status == PromptStatus.OK)
                    {
                        // Support ?
                        //
                        if (res.ObjectId.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Support))))
                        {
                            // Yes. Take it
                            //
                            id = res.ObjectId;
                            break;
                        }
                    }
                    else
                    if (res.Status == PromptStatus.Cancel)
                    {
                        return;
                    }
                }

                // Find parameters
                //
                ParameterList plist = SupportHelper.GetSupportParameters (id);
                if (plist == null || plist.Count == 0)
                {
                    ed.WriteMessage("\nSupport has no parameters.");
                    return;
                }

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
                                if (!SupportHelper.UpdateSupportEntity (id, plist))
                                {
                                    ed.WriteMessage("\nSupport update failed.");
                                }
                            }
                            break;
                        }
                    }
                }
            }
            catch (System.Exception)
            {
            }
        }


        // Register custom script
        //
        [CommandMethod("SupportRegister")]
        public static void SupportRegister()
        {
            try
            {
                Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;

                // Ask for a script name
                //
                PromptStringOptions sopts = new PromptStringOptions("\nScript name:");
                while(true)
                {
                    PromptResult sres = ed.GetString(sopts);
                    if (sres.Status != PromptStatus.OK)
                    {
                        return;
                    }
                    
                    // Exist?
                    //
                    SupportInfo info = SupportHelper.FindSupportInfo(sres.StringResult);
                    if (info == null)
                    {
                        ed.WriteMessage("\nScript doesn't exist.");
                        continue;
                    }

                    // Type
                    //
                    PromptKeywordOptions opts = new PromptKeywordOptions("\nScript type:");
                    opts.Keywords.Add("Simple");
                    opts.Keywords.Add("Dummyleg");
                    opts.Keywords.Add("Trapezebar");
                    opts.Keywords.Add("Vertical");
                    opts.Keywords.Add("Base");
                    switch (info.Type)
                    {
                    case SupportType.eSupport:
                    default:
                        opts.Keywords.Default = "Simple";
                        break;
                    case SupportType.eDummyLeg:
                        opts.Keywords.Default = "Dummyleg";
                        break;
                    case SupportType.eTrapezeBar:
                        opts.Keywords.Default = "Trapezebar";
                        break;
                    case SupportType.eVerticalSupport:
                        opts.Keywords.Default = "Vertical";
                        break;
                    case SupportType.eBaseSupport:
                        opts.Keywords.Default = "Base";
                        break;
                    }
                    PromptResult res = ed.GetKeywords(opts);
                    if (res.Status == PromptStatus.OK)
                    {
                        if (res.StringResult == "Simple")
                        {
                            info.Type = SupportType.eSupport;
                        }
                        else
                        if (res.StringResult == "Dummyleg")
                        {
                            info.Type = SupportType.eDummyLeg;
                        }
                        else
                        if (res.StringResult == "Trapezebar")
                        {
                            info.Type = SupportType.eTrapezeBar;
                        }
                        else
                        if (res.StringResult == "Vertical")
                        {
                            info.Type = SupportType.eVerticalSupport;
                        }
                        else
                        if (res.StringResult == "Base")
                        {
                            info.Type = SupportType.eBaseSupport;
                        }
                    }

                    // Register
                    //
                    info.ToRegistry(SupportHelper.GetConfigurationSection(true));
                    break;
                }
            }
            catch (System.Exception)
            {
            }
        }


        protected static void PlaceSupport (Support support, PartSizeProperties part)
        {
            try
            {
                Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;
                Database db                 = AcadApp.DocumentManager.MdiActiveDocument.Database;
                Project currentProject      = PlantApp.CurrentProject.ProjectParts["Piping"];
                DataLinksManager dlm        = currentProject.DataLinksManager;
                DataLinksManager3d dlm3d    = DataLinksManager3d.Get3dManager(dlm);

                while (true)
                {
                    // Drag. Ask for a placement point
                    //
                    SuppDragPoint dragPoint = new SuppDragPoint(support, part);
                    dragPoint.StartPointMonitor();
                    PromptStatus st = ed.Drag(dragPoint).Status;
                    dragPoint.StopPointMonitor();
                    if (st != PromptStatus.OK)
                    {
                        return;
                    }

                    if (dragPoint.SnapId.IsNull)
                    {
                        // No snapping. Ask for a rotation
                        //
                        SuppDragAngle dragAngle = new SuppDragAngle(support);
                        ed.Drag(dragAngle);
                    }

                    // Make a copy to place
                    //
                    Support supportCopy = (Support)support.Clone();

                    // Add
                    //
                    ObjectId supId;
                    using (Transaction tr = db.TransactionManager.StartOpenCloseTransaction())
                    {
                        using (PipingObjectAdder pipeObjAdder = new PipingObjectAdder(dlm3d, db))
                        {
                            pipeObjAdder.Add(part, supportCopy);
                            tr.AddNewlyCreatedDBObject(supportCopy, true);
                            supId = supportCopy.ObjectId;
                        }

                        tr.Commit();
                    }

                    if (!dragPoint.SnapId.IsNull)
                    {
                        // Connect
                        //
                        SupportHelper.ConnectSupport(supId, dragPoint.SnapId);

                        // Cut if needed
                        //
                        SupportHelper.CutSupport(supId, dragPoint.SnapId);
                    }
                }
            }
            catch (System.Exception)
            {
            }
        }


        // Dragger classes
        //
        public class SuppDragPoint : EntityJig
        {
            PartSizeProperties m_Part;
            Point3d m_CurrentPos;
            bool m_bNewSnap;
            JigPromptPointOptions m_Opts;

            public ObjectId SnapId
            {
                get;
                set;
            }

            public Autodesk.ProcessPower.PnP3dObjects.Port SnapPort
            {
                get;
                set;
            }
            
            public SuppDragPoint(Support support, PartSizeProperties part)
                : base(support)
            {
                m_Part = part;
                m_CurrentPos = support.Position;
                m_bNewSnap = false;

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
                // Do we have osnapped ent?
                //
                if (!SnapId.IsNull)
                {
                    // Yes. New entity?
                    //
                    if (m_bNewSnap)
                    {
                        // Size of snapped ent
                        //
                        try
                        {
                            using (Transaction tr = AcadApp.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartOpenCloseTransaction())
                            {
                                Autodesk.ProcessPower.PnP3dObjects.Part part = tr.GetObject(SnapId, OpenMode.ForRead) as Autodesk.ProcessPower.PnP3dObjects.Part;
                                PartSizeProperties partProps = part.PartSizeProperties;
                                if (!m_Part.NominalDiameter.Equals(partProps.NominalDiameter))
                                {
                                    // Resize
                                    //
                                    if (!SupportHelper.ResizeSupport((Support)Entity, m_Part, partProps.NominalDiameter))
                                    {
                                        // Clear snap
                                        //
                                        SnapId = ObjectId.Null;
                                    }
                                }
 
                                tr.Commit();
                            }
                        }
                        catch (System.Exception)
                        {
                        }
                    }
                }

                // Still have osnap after resizing?
                //
                if (!SnapId.IsNull)
                {
                    // Align
                    //
                    if (!SupportHelper.AlignSupport((Support)Entity, m_Part, false, SnapId, SnapPort))
                    {
                        // Clear snap
                        //
                        SnapId = ObjectId.Null;
                    }
                }

                if (SnapId.IsNull)
                {
                    // No osnap. Just set new point
                    //
                    ((Support)Entity).Position = m_CurrentPos;
                }

                return true;
            }

            public void StartPointMonitor()
            {
                AcadApp.DocumentManager.MdiActiveDocument.Editor.PointMonitor += new PointMonitorEventHandler(drag_PointMonitorEventHandler);
            }

            public void StopPointMonitor()
            {
                AcadApp.DocumentManager.MdiActiveDocument.Editor.PointMonitor -= new PointMonitorEventHandler(drag_PointMonitorEventHandler);
            }

            public void drag_PointMonitorEventHandler(object sender, PointMonitorEventArgs args)
            {
                InputPointContext context = args.Context;
                if (context == null)
                {
                    return;
                }
                if (!context.PointComputed)
                {
                    return;
                }

                // Clear current snap
                //
                m_bNewSnap = false;
                ObjectId oldSnapId = SnapId;
                SnapId = ObjectId.Null;

                // Check snapping
                //
                if (context.ObjectSnapMask == 0)
                {
                    // No snap
                    //
                    return;
                }
                Point3d snapPoint = context.ObjectSnappedPoint;
                FullSubentityPath []paths = context.GetKeyPointEntities();
                if (paths == null)
                {
                    return;
                }
                foreach (FullSubentityPath path in paths)
                {
                    ObjectId []ids = path.GetObjectIds();
                    if (ids.Length == 0)
                    {
                        continue;
                    }
                    ObjectId id = ids[0];

                    // Can snap to pipe or fitting
                    //
                    if (!id.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Pipe))) &&
                        !id.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(PipeInlineAsset))))
                    {
                        continue;
                    }

                    // Can snap ?
                    //
                    Autodesk.ProcessPower.PnP3dObjects.Port port = null;
                    if (SupportHelper.FindSupportAlignment(m_Part, id, snapPoint, out port))
                    {
                        // Take it
                        //
                        SnapId = id;
                        SnapPort = port;
                        if (id != oldSnapId)
                        {
                            m_bNewSnap = true;
                        }
                        break;
                    }
                }
            }
        }

        public class SuppDragAngle : EntityJig
        {
            Double m_CurrentAngle;
            Double m_DefaultAngle;
            Point3d m_Position;
            Vector3d m_ZAxis;
            Matrix3d m_Trans;
            JigPromptAngleOptions m_Opts;

            public SuppDragAngle(Support support)
                : base(support)
            {
                m_CurrentAngle = m_DefaultAngle = support.Rotation;
                m_Position = support.Position;
                m_ZAxis = support.ZAxis;
                m_Trans = Matrix3d.Identity;

                m_Opts = new JigPromptAngleOptions("\nSelect rotation angle");
                m_Opts.UserInputControls = UserInputControls.NullResponseAccepted;
                m_Opts.Cursor |= CursorType.RubberBand;
                m_Opts.BasePoint = support.Position;
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
                            m_Trans = Matrix3d.Rotation(res.Value - m_CurrentAngle, m_ZAxis, m_Position);
                            m_CurrentAngle = res.Value;
                        }
                        break;

                    default:
                        // Restore default
                        //
                        m_Trans = Matrix3d.Rotation(m_DefaultAngle - m_CurrentAngle, m_ZAxis, m_Position);
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

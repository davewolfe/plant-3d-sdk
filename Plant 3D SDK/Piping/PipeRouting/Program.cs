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
using Autodesk.ProcessPower.PnP3dPipeRouting;
using Autodesk.ProcessPower.P3dUI;
using Autodesk.ProcessPower.PnP3dTagUtil;
using Autodesk.ProcessPower.ACPUtils;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using PlantApp = Autodesk.ProcessPower.PlantInstance.PlantApplication;

[assembly: Autodesk.AutoCAD.Runtime.ExtensionApplication(null)]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(PipeRoutingSample.Program))]

namespace PipeRoutingSample
{
    public class Program
    {
        // Route pipes
        //
        [CommandMethod("PipeRoute")]
        public static void PipeRoute()
        {
            try
            {
                using (DoRoute route = new DoRoute())
                {
                    route.Run();
                }
            }
            catch (System.Exception)
            {
            }
        }

        // Place fitting
        //
        [CommandMethod("FittingAdd")]
        public static void FittingAdd()
        {
            try
            {
                PlaceFitting place = new PlaceFitting();
                place.Run();
            }
            catch (System.Exception)
            {
            }
        }

        // Edit fitting
        //
        [CommandMethod("FittingEdit")]
        public static void FittingEdit()
        {
            Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                // Select part
                //
                ObjectId id = ObjectId.Null;
                while (true)
                {
                    PromptEntityResult res = ed.GetEntity("\nSelect parametric part: ");
                    if (res.Status == PromptStatus.OK)
                    {
                        // Fitting, equipment and support
                        //
                        if (res.ObjectId.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(PipeInlineAsset))) ||
                            res.ObjectId.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Support))) ||
                            res.ObjectId.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Equipment))))
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
                ParametricPart param = null;
                try
                {
                    param = new ParametricPart(id);
                }
                catch (System.Exception ex)
                {
                    if (ex.Message != null)
                    {
                        ed.WriteMessage(ex.Message);
                    }
                    ed.WriteMessage("\nPart has no parameters.");
                    return;
                }

                ParameterList plist = param.Parameters;
                if (plist == null || plist.Count == 0)
                {
                    ed.WriteMessage("\nPart has no parameters.");
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
                                if (!param.UpdateEntity (id, plist))
                                {
                                    ed.WriteMessage("\nFitting update failed.");
                                }
                            }
                            break;
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                if (ex.Message != null)
                {
                    ed.WriteMessage(ex.Message);
                }
            }
        }

        // Substitute fitting
        //
        [CommandMethod("FittingSubstitute")]
        public static void FittingSubstitute()
        {
            Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                // Select fitting
                //
                ObjectId id = ObjectId.Null;
                while (true)
                {
                    PromptEntityResult res = ed.GetEntity("\nSelect fitting: ");
                    if (res.Status == PromptStatus.OK)
                    {
                        // Fitting, equipment and support
                        //
                        if (res.ObjectId.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(PipeInlineAsset))))
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

                // Properties
                //
                PartSizeProperties props = null;
                using (Transaction tr = id.Database.TransactionManager.StartOpenCloseTransaction())
                {
                    PipeInlineAsset fitting = tr.GetObject(id, OpenMode.ForRead) as PipeInlineAsset;
                    props = fitting.PartSizeProperties;
                    tr.Commit();
                }

                // Select all the parts of that type and size
                //
                List<PartSizeProperties> parts = new List<PartSizeProperties>();
                SpecManager specMgr = SpecManager.GetSpecManager();
                SpecPartReader reader = specMgr.SelectParts(props.Spec, props.Type, props.NominalDiameter);
                while (reader.Next())
                {
                    // Take, if not current part
                    //
                    if (String.Compare(props.PropValue("SizeRecordId").ToString(), reader.Current.PropValue("SizeRecordId").ToString(), true) != 0)
                    {
                        SpecPart part = new SpecPart();
                        part.AssignFrom(reader.Current);
                        parts.Add(part);
                    }
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
                            s = s.Replace(",", "");
                        }
                    }
                    catch (System.Exception)
                    {
                    }
                    kwds.Add(cnt.ToString() + "." + s);
                }
                if (kwds.Count == 0)
                {
                    ed.WriteMessage("\nNo parts to substitute with found");
                    return;
                }
                PromptResult res1 = ed.GetKeywords("\nSelect part", kwds.ToArray());
                if (res1.Status != PromptStatus.OK)
                {
                    return;
                }
                PartSizeProperties newProps = parts[kwds.IndexOf(res1.StringResult)];

                // Create fitting
                //
                ContentManager conMgr = ContentManager.GetContentManager();
                ObjectId symbolId = conMgr.GetSymbol(newProps, id.Database, props.LinearUnit);
                PipeInlineAsset newFitting = new PipeInlineAsset();
                newFitting.SymbolId = symbolId;
                Autodesk.ProcessPower.PnP3dObjects.PortCollection newPorts = newFitting.GetPorts(PortType.Static);

                // Substitue
                //
                RoutingHelper.SubstituteAsset(id, newProps, newPorts, symbolId);
            }
            catch (System.Exception ex)
            {
                if (ex.Message != null)
                {
                    ed.WriteMessage(ex.Message);
                }
            }
        }
    }

    public class DoRoute : DrawJig, IDisposable
    {
        // Current settings
        //
        UISettings m_Settings = new UISettings();
        UISettings.SettingType m_ChangedSetting = UISettings.SettingType.PipingLowerBound;
        String m_Spec = null;
        NominalDiameter m_Size;
        bool m_bCutbackElbow = false;
        bool m_bBentPipe = false;
        bool m_bStubin = false;
        bool m_bToleranceRouting = false;
        String m_LineNumber;

        // Internals
        //
        int m_GroupId = -1;
        bool m_bCanCutbackPipe = false;
        bool m_bCanContinuePipe= false;
        double m_MaxCutbackLength = 0.0;

        // Snapping
        //
        Pair m_SnapPair = null;
        PartSizeProperties m_SnapPart = null;

        // Last created part (or the first selected for the first point)
        //
        Pair m_LastPair = null;
        PartSizeProperties m_LastPart = null;

        // Jig
        //
        JigPromptPointOptions m_Opts = null;
        Point3d m_CurrentPos;

        // Current pipe and elbow parts
        //
        PipePart m_CurrentPipe = null;
        ElbowPart[] m_ElbowParts = null;
        ElbowPart m_CurrentElbow = null;

        // Current parts
        //
        BranchWrapper m_Branch = null;
        ConnectorWrapper m_ReducerConnector = null;
        ReducerWrapper m_Reducer = null; 
        ConnectorWrapper m_ElbowConnector = null;
        ElbowWrapper m_Elbow = null;
        ConnectorWrapper m_PipeConnector = null;
        PipeWrapper m_Pipe = null;

        // Constants
        //
        double m_ZeroAngle = 1e-10;
        double m_MinElbowAngle = 3.14159/60.0;       // 3 degrees
        double m_ZeroDist = 1e-10;

        public DoRoute () : base()
        {
        }

        public void Dispose()
        {
            Clear(true);
        }

        private void Clear(bool bClearBranch)
        {
            if (bClearBranch)
            {
                if (m_Branch != null)
                {
                    m_Branch.Dispose();
                    m_Branch = null;
                }
            }
            if (m_ReducerConnector != null)
            {
                m_ReducerConnector.Dispose();
                m_ReducerConnector = null;
            }
            if (m_Reducer != null)
            {
                m_Reducer.Dispose();
                m_Reducer = null;
            }
            if (m_ElbowConnector != null)
            {
                m_ElbowConnector.Dispose();
                m_ElbowConnector = null;
            }
            if (m_Elbow != null)
            {
                m_Elbow.Dispose();
                m_Elbow = null;
            }
            if (m_PipeConnector != null)
            {
                m_PipeConnector.Dispose();
                m_PipeConnector = null;
            }
            if (m_Pipe != null)
            {
                m_Pipe.Dispose();
                m_Pipe = null;
            }
        }

        public void Run()
        {
            Editor ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                // Current elevation
                //
                double elevation = m_Settings.Elevation;
                AcadApp.SetSystemVariable("Elevation", elevation);

                // Ask for a start point
                //
                if (!GetStartPoint())
                {
                    return;
                }

                // Current settings
                //
                m_Spec = m_Settings.CurrentSpec;
                m_Size = NominalDiameter.FromDisplayString(null, m_Settings.CurrentSize);
                m_bCutbackElbow = m_Settings.CutbackFlag;
                m_bBentPipe = m_Settings.PipeBendFlag;
                m_bStubin = m_Settings.StubinFlag;
                m_bToleranceRouting = m_Settings.CompassToleranceRoutingEnabled;
                if (m_GroupId < 0)
                {
                    m_LineNumber = m_Settings.CurrentLineNumber;
                }

                // Next points
                //
                m_Opts = new JigPromptPointOptions("\nSelect point");
                m_Opts.UserInputControls = (UserInputControls.Accept3dCoordinates | UserInputControls.NullResponseAccepted);
                m_Opts.Keywords.Add("Size");
                m_Opts.Keywords.Add("SPecification");
                m_Opts.Keywords.Add("Cutbackelbow");
                m_Opts.Keywords.Add("pipeBEnd");
                m_Opts.Keywords.Add("STub-in");
                while(true)
                {
                    // Create current parts
                    //
                    CreateParts(true);

                    // Jig with moint monitor and UISetting reactor
                    //
                    m_Settings.OnSettingChanged += new UISettings.SettingChangedEventHandler(Settings_OnSettingChanged);
                    StartPointMonitor();
                    PromptResult res = ed.Drag(this);
                    StopPointMonitor();
                    m_Settings.OnSettingChanged -= new UISettings.SettingChangedEventHandler(Settings_OnSettingChanged);

                    // Any settings changed ?
                    //
                    if (m_ChangedSetting != UISettings.SettingType.PipingLowerBound)
                    {
                        switch (m_ChangedSetting)
                        {
                            case UISettings.SettingType.Spec:
                                m_Spec = m_Settings.CurrentSpec;
                                m_CurrentPipe = null;
                                m_ElbowParts = null;
                                m_CurrentElbow = null;
                                break;

                            case UISettings.SettingType.Size:
                                m_Size = NominalDiameter.FromDisplayString(null, m_Settings.CurrentSize);
                                m_CurrentPipe = null;
                                m_ElbowParts = null;
                                m_CurrentElbow = null;
                                break;

                            case UISettings.SettingType.Cutback:
                                m_bCutbackElbow = m_Settings.CutbackFlag;
                                break;

                            case UISettings.SettingType.PipeBend:
                                m_bBentPipe = m_Settings.PipeBendFlag;
                                break;

                            case UISettings.SettingType.Stubin:
                                m_bStubin = m_Settings.StubinFlag;
                                break;

                            case UISettings.SettingType.CompassToleranceRoutingEnabled:
                                m_bToleranceRouting = m_Settings.CompassToleranceRoutingEnabled;
                                break;

                            case UISettings.SettingType.LineNumber:
                                m_LineNumber = m_Settings.CurrentLineNumber;
                                m_GroupId = -1;
                                break;
                        }

                        // Clear and restart routing
                        //
                        m_ChangedSetting = UISettings.SettingType.PipingLowerBound;
                        continue;
                    }

                    switch (res.Status)
                    {
                        case PromptStatus.None:
                        case PromptStatus.Cancel:
                            return;

                        case PromptStatus.OK:
                            if (m_SnapPair != null)
                            {
                                // Try auto routing
                                //
                                bool bCancelled;
                                if (DoAutoRoute(out bCancelled))
                                {
                                    // Finished routing
                                    //
                                    return;
                                }
                                else
                                if (bCancelled)
                                {
                                    break;
                                }
                            }
                            AppendParts();
                            break;

                        case PromptStatus.Keyword:
                            if (res.StringResult == "Size")
                            {
                                PromptResult sres = ed.GetString("\nNew size: ");
                                if (sres.Status == PromptStatus.OK)
                                {
                                    m_Settings.CurrentSize = sres.StringResult;
                                    m_Size = NominalDiameter.FromDisplayString(null, m_Settings.CurrentSize);
                                    m_CurrentPipe = null;
                                    m_ElbowParts = null;
                                    m_CurrentElbow = null;
                                }
                            }
                            else
                            if (res.StringResult == "SPecification")
                            {
                                PromptResult sres = ed.GetString("\nNew spec name: ");
                                if (sres.Status == PromptStatus.OK)
                                {
                                    m_Settings.CurrentSpec = sres.StringResult;
                                    m_Spec = m_Settings.CurrentSpec;
                                    m_CurrentPipe = null;
                                    m_ElbowParts = null;
                                    m_CurrentElbow = null;
                                }
                            }
                            else
                            if (res.StringResult == "Cutbackelbow")
                            {
                                m_bCutbackElbow = !m_bCutbackElbow;
                                m_Settings.CutbackFlag = m_bCutbackElbow; 
                            }
                            else
                            if (res.StringResult == "pipeBEnd")
                            {
                                m_bBentPipe = !m_bBentPipe;
                                m_Settings.PipeBendFlag = m_bBentPipe;
                            }
                            else
                            if (res.StringResult == "STub-in")
                            {
                                m_bStubin = !m_bStubin;
                                m_Settings.StubinFlag = m_bStubin;
                            }
                            break;

                        case PromptStatus.Modeless:
                            // Never comes
                            //
                            break;

                        default:
                            break;
                    }
                }
            }
            catch (System.Exception ex)
            {
                if (ex.Message != null)
                {
                    ed.WriteMessage(ex.Message);
                }
            }
        }

        public bool GetStartPoint()
        {
            try
            {
                Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;

                // Ask for a point
                //
                StartPointMonitor();
                PromptPointResult res = ed.GetPoint("\nSelect start point: ");
                StopPointMonitor();
                if (res.Status != PromptStatus.OK)
                {
                    return false;
                }

                // Do we have snap ?
                //
                if (m_SnapPair != null)
                {
                    // Take it
                    //
                    m_LastPair = m_SnapPair;
                    m_LastPart = m_SnapPart;

                    // Line group from the part
                    //
                    m_GroupId = PnP3dTagFormat.lineGroupIdFromObjId(m_SnapPair.ObjectId);
                    if (m_GroupId > 0)
                    {
                        m_LineNumber = PnP3dTagFormat.lineNumberTagFromGroup(m_GroupId);
                        if (!String.IsNullOrEmpty(m_LineNumber))
                        {
                            m_Settings.CurrentLineNumber = m_LineNumber;
                        }
                    }
                }
                else
                {
                    // Create a pair with only a point
                    //
                    Autodesk.ProcessPower.PnP3dObjects.Port port = new Autodesk.ProcessPower.PnP3dObjects.Port();
                    port.Position = res.Value; 
                    m_LastPair = new Pair();
                    m_LastPair.Port = port;
                    m_LastPart = null;
                }

                return true;
            }
            catch (System.Exception)
            {
            }

            return false;
        }

        void Settings_OnSettingChanged(object sender, UISettings.SettingType type)
        {
            // Take supported flags
            //
            switch (type)
            {
                case UISettings.SettingType.Spec:
                case UISettings.SettingType.Size:
                case UISettings.SettingType.Cutback:
                case UISettings.SettingType.PipeBend:
                case UISettings.SettingType.Stubin:
                case UISettings.SettingType.CompassToleranceRoutingEnabled:
                case UISettings.SettingType.LineNumber:
                    m_ChangedSetting = type;
                    break;
            }
        }

        public void CreateParts(bool bCreateBranch)
        {
            Clear(bCreateBranch);
            m_bCanCutbackPipe = false;
            m_bCanContinuePipe = false;
            m_MaxCutbackLength = 0.0;

            Database db                 = AcadApp.DocumentManager.MdiActiveDocument.Database;              
            Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;

            PartSizeProperties lastPart = m_LastPart;
            Autodesk.ProcessPower.PnP3dObjects.Port lastPort = m_LastPair.Port;

            if (bCreateBranch)
            {
                // Recerate branch if needed
                // 
                if (!m_LastPair.ObjectId.IsNull)
                {
                    if (String.IsNullOrEmpty(m_LastPair.Port.Name))
                    {
                        // We need a branch
                        //
                        try
                        {
                            m_Branch = new BranchWrapper(m_LastPair.ObjectId, m_LastPair.Port, m_Size, m_bStubin);
                            lastPart = m_Branch.BranchProperties; 
                            lastPort = m_Branch.BranchEndPort;
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage("\nCannot create branch");
                            if (ex.Message != null)
                            {
                                ed.WriteMessage(ex.Message);
                            }
                        }
                    }
                }
            }
            else
            {
                // Use existing branch
                //
                lastPart = m_Branch.BranchProperties; 
                lastPort = m_Branch.BranchEndPort;
            }

            if (lastPart != null)
            {
                // Do we need a reducer?
                //
                if (!lastPart.NominalDiameter.Equals(m_Size))
                {
                    // Find it
                    //
                    try
                    {
                        m_Reducer = new ReducerWrapper(m_Spec, lastPart.NominalDiameter, m_Size, null, null, null);

                        // Connect
                        //
                        m_ReducerConnector = new ConnectorWrapper();
                        m_ReducerConnector.Connect(lastPart, lastPort, m_Reducer.PartSizeProperties, m_Reducer.StartPort, null);

                        // Align geometrically
                        //
                        Matrix3d mat = RoutingHelper.CalculateAttachMatrix(m_ReducerConnector.StartPort,
                                                                           RoutingHelper.CalculateNormal(m_ReducerConnector.StartPort, null),
                                                                           lastPort,
                                                                           RoutingHelper.CalculateNormal(lastPort, null));
                        m_ReducerConnector.TransformBy(mat);
                        mat = RoutingHelper.CalculateAttachMatrix(m_Reducer.StartPort,
                                                                  RoutingHelper.CalculateNormal(m_Reducer.StartPort, null),
                                                                  m_ReducerConnector.EndPort,
                                                                  RoutingHelper.CalculateNormal(m_ReducerConnector.EndPort, null));
                        m_Reducer.TransformBy(mat);

                        lastPart = m_Reducer.PartSizeProperties; 
                        lastPort = m_Reducer.EndPort;
                    }
                    catch (System.Exception ex)
                    {
                        m_Reducer = null;
                        m_ReducerConnector = null;
                        if (ex.Message != null)
                        {
                            ed.WriteMessage(ex.Message);
                        }
                    }
                }
            }

            // Create a pipe
            //
            try
            {
                if (m_CurrentPipe == null)
                {
                    // Find it
                    //
                    m_CurrentPipe = new PipePart(m_Size, m_Spec);
                }

                // Create pipe wrapper
                //
                m_Pipe = m_CurrentPipe.CreatePipe();

                // Make short pipe
                //
                m_Pipe.StartPoint = lastPort.Position;
                Vector3d dir = lastPort.Direction;
                if (dir.IsZeroLength())
                {
                    dir = Vector3d.XAxis;   //something
                }
                m_Pipe.EndPoint = m_Pipe.StartPoint + 2 * m_Pipe.MinLength * dir;
            }
            catch (System.Exception ex)
            {
                m_Pipe = null;
                if (ex.Message != null)
                {
                    ed.WriteMessage(ex.Message);
                }
            }

            // Load elbows, in case we need them
            //
            if (m_ElbowParts == null)
            {
                m_ElbowParts = ElbowPart.FindElbows(m_Size, m_Spec, false);
            }

            // And if we don't have branch or reducer and last entity is a pipe,
            // check that we can cut it back with an elbow or continue with current pipe 
            //
            if (m_Pipe != null &&
                m_Branch == null && m_Reducer == null &&
                m_LastPart != null &&
                !m_LastPair.ObjectId.IsNull &&
                !String.IsNullOrEmpty(m_LastPair.Port.Name) &&
                m_LastPair.ObjectId.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Pipe))))
            {
                // Pipe must be writable
                //
                using (Transaction tr = db.TransactionManager.StartOpenCloseTransaction())
                {
                    try
                    {
                        // Open and get pipe length
                        //
                        Pipe pipe = tr.GetObject(m_LastPair.ObjectId, OpenMode.ForWrite) as Pipe;
                        bool bFixedLength = pipe.FixedLength;
                        double length = pipe.Length;

                        // Min/Max length
                        // 
                        PipePart pipePart = new PipePart(m_LastPart, ObjectId.Null);
                        double minLength = pipePart.MinLength;
                        double maxLength = 0.0;
                        if (bFixedLength)
                        {
                            maxLength = pipePart.MaxLength;
                            if (maxLength <= 0.0)
                            {
                                bFixedLength = false;
                            }
                        }
                        
                        // Can cutback, if the pipe is longer than min,
                        // and not at max length if FLP
                        //
                        if (length > minLength &&
                            (!bFixedLength || Math.Abs(maxLength-length) < m_ZeroDist))
                        {
                            m_bCanCutbackPipe = true;
                            m_MaxCutbackLength = length - minLength;
                        }

                        // Can continue if not FLP, or shorter than max
                        //
                        if (!bFixedLength || length < maxLength)
                        {
                            // And the parts must be identical
                            //
                            if (String.Compare(m_LastPart.Spec, m_Pipe.PartSizeProperties.Spec, true) == 0 &&
                                String.Compare(m_LastPart.PropValue("SizeRecordId").ToString(), m_Pipe.PartSizeProperties.PropValue("SizeRecordId").ToString(), true) == 0)
                            {
                                m_bCanContinuePipe = true;
                                m_Pipe.StartLength = length;
                            }
                        }
                    }
                    catch(System.Exception)
                    {
                    }

                    tr.Commit();
                }
            }
        }

        public void AppendParts()
        {
            Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;
            Database db                 = AcadApp.DocumentManager.MdiActiveDocument.Database;              
            
            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Line group
                    // TODO: check locked group
                    //
                    if (m_GroupId < 0)
                    {
                        m_GroupId = PnP3dTagFormat.findOrCreateNewLineGroup(m_LineNumber);
                    }

                    Pair lastPair = new Pair();
                    lastPair.ObjectId = m_LastPair.ObjectId;
                    lastPair.Port = m_LastPair.Port;
                    PartSizeProperties lastPart = m_LastPart;

                    // Append parts
                    //
                    if (m_Branch != null)
                    {
                        m_Branch.Append(m_GroupId);
                        lastPair.ObjectId = m_Branch.LastCreatedId;
                        lastPair.Port = m_Branch.BranchEndPort;
                        lastPart = m_Branch.BranchProperties; 
                    }

                    if (m_Reducer != null)
                    {
                        Autodesk.ProcessPower.PnP3dObjects.Port startPort = m_Reducer.StartPort;
                        Autodesk.ProcessPower.PnP3dObjects.Port endPort = m_Reducer.EndPort;
                        ObjectId id = m_Reducer.Append(m_GroupId);

                        if (m_ReducerConnector != null)
                        {
                            Pair startPair = new Pair();
                            startPair.ObjectId = id;
                            startPair.Port = startPort;
                            m_ReducerConnector.Append(lastPair, startPair, m_GroupId);
                        }

                        lastPair.ObjectId = id;
                        lastPair.Port = endPort;
                        lastPart = m_Reducer.PartSizeProperties; 
                    }

                    if (m_Elbow != null)
                    {
                        Autodesk.ProcessPower.PnP3dObjects.PortCollection ports = m_Elbow.GetPorts(PortType.Static);
                        ObjectId id = m_Elbow.Append(m_GroupId);

                        if (m_ElbowConnector != null)
                        {
                            // If we need to cutback last pipe, Append will do that
                            //
                            Pair startPair = new Pair();
                            startPair.ObjectId = id;
                            startPair.Port = ports[0];
                            m_ElbowConnector.Append(lastPair, startPair, m_GroupId);
                        }

                        lastPair.ObjectId = id;
                        lastPair.Port = ports[1];
                        lastPart = m_Elbow.PartSizeProperties;
                    }

                    if (m_Pipe != null)
                    {
                        if (m_PipeConnector != null)
                        {
                            // Append unconnected pipe and connect it with the connecter
                            //
                            m_Pipe.Append(null, null, m_GroupId);
                            m_PipeConnector.Append(lastPair, m_Pipe.FirstCreatedPort, m_GroupId);
                        }
                        else
                        {
                            // Continue previous pipe
                            //
                            System.Diagnostics.Debug.Assert(m_Elbow == null);
                            m_Pipe.Append(lastPair, null, m_GroupId);
                        }

                        lastPair = m_Pipe.LastCreatedPort;
                        lastPart = m_Pipe.PartSizeProperties; 
                    }

                    tr.Commit();

                    // Store last pair
                    //
                    m_LastPair = lastPair;
                    m_LastPart = lastPart;
                }
            }
            catch (System.Exception ex)
            {
                if (ex.Message != null)
                {
                    ed.WriteMessage(ex.Message);
                }
            }
        }

        public void StartPointMonitor()
        {
            m_SnapPair = null;
            m_SnapPart = null;
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
            m_SnapPair = null;
            m_SnapPart = null;

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

            Database db                 = AcadApp.DocumentManager.MdiActiveDocument.Database;              
            ConnectionManager conMan    = new ConnectionManager();
            Project currentProject      = PlantApp.CurrentProject.ProjectParts["Piping"];
            DataLinksManager dlm        = currentProject.DataLinksManager;

            // Find part to snap to
            //
            foreach (FullSubentityPath path in paths)
            {
                ObjectId []ids = path.GetObjectIds();
                if (ids.Length == 0)
                {
                    continue;
                }
                ObjectId id = ids[0];

                // Can snap to pipe, fitting and equipment
                //
                if (!id.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Pipe))) &&
                    !id.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(PipeInlineAsset))) &&
                    !id.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Equipment))))
                {
                    continue;
                }

                // Get ports and properties
                //
                Autodesk.ProcessPower.PnP3dObjects.PortCollection ports = null;
                PartSizeProperties props = null;
                bool bEditable = false;
                Transaction tr = null;
                try
                {
                    tr = db.TransactionManager.StartOpenCloseTransaction();
                    Autodesk.ProcessPower.PnP3dObjects.Part part = tr.GetObject(id, OpenMode.ForRead) as Autodesk.ProcessPower.PnP3dObjects.Part;
                    ports = part.GetPorts(PortType.Static);
                    props = part.PartSizeProperties;

                    // And check that the part is editable (not from xref and can be open for write) 
                    //
                    if (part.OwnerId == db.CurrentSpaceId)
                    {
                        tr.GetObject(id, OpenMode.ForWrite);
                        bEditable = true;
                    }
                }
                catch (System.Exception)
                {
                }
                finally
                {
                    if (tr != null)
                    {
                        tr.Commit();
                    }
                }

                if (props == null)
                {
                    continue;
                }

                // Find port
                //
                if (ports != null)
                {
                    foreach (Autodesk.ProcessPower.PnP3dObjects.Port port in ports)
                    {
                        if (port.Position != snapPoint)
                        {
                            continue;
                        }

                        Pair pair = new Pair();
                        pair.ObjectId = id;
                        pair.Port = port;

                        // Connected ?
                        //
                        if (!conMan.IsConnected(pair))
                        {
                            // Take this port
                            //
                            m_SnapPair = pair;
                            m_SnapPart = props;
                            return;
                        }

                        // And we can branch connected ports for elbows and tees,
                        // replacing them with the tee and cross
                        //
                        if (bEditable &&
                            id.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(PipeInlineAsset))))
                        {
                            String className = dlm.GetObjectClassname(id);
                            if (dlm.HasTable(className))
                            {
                                PnPTable tbl = dlm.GetPnPDatabase().Tables[className];
                                if (tbl.IsKindOf("Elbow") ||
                                    tbl.IsKindOf("SingleBranchFitting"))
                                {
                                    // Take this port without port name, what would mean branching
                                    //
                                    pair.Port.Name = String.Empty;
                                    m_SnapPair = pair;
                                    m_SnapPart = props;
                                    return;
                                }
                            }
                        }

                        // Can't snap to this port
                        //
                        return;
                    }
                }

                // We don't snap to the port
                // But can branch pipe and fitting
                //
                if (!bEditable &&
                    !id.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Pipe))) &&
                    !id.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(PipeInlineAsset))))
                {
                    continue;
                }

                // Branch
                //
                Autodesk.ProcessPower.PnP3dObjects.Port fakePort = new Autodesk.ProcessPower.PnP3dObjects.Port();
                fakePort.Position = snapPoint; 
                m_SnapPair = new Pair();
                m_SnapPair.ObjectId = id;
                m_SnapPair.Port = fakePort;
                m_SnapPart = props;
                return;
            }
        }

        protected override bool WorldDraw(Autodesk.AutoCAD.GraphicsInterface.WorldDraw draw)
        {
            bool ret = true;

            if (m_Branch != null)
            {
                ret &= m_Branch.WorldDraw(draw);
            }
            if (m_ReducerConnector != null)
            {
                ret &= m_ReducerConnector.WorldDraw(draw);
            }
            if (m_Reducer != null)
            {
                ret &= m_Reducer.WorldDraw(draw);
            }
            if (m_ElbowConnector != null)
            {
                ret &= m_ElbowConnector.WorldDraw(draw);
            }
            if (m_Elbow != null)
            {
                ret &= m_Elbow.WorldDraw(draw);
            }
            if (m_PipeConnector != null)
            {
                ret &= m_PipeConnector.WorldDraw(draw);
            }
            if (m_Pipe != null)
            {
                ret &= m_Pipe.WorldDraw(draw);
            }

            return ret;
        }

        protected override void ViewportDraw(Autodesk.AutoCAD.GraphicsInterface.ViewportDraw draw)
        {
            if (m_Branch != null)
            {
                m_Branch.ViewportDraw(draw);
            }
            if (m_ReducerConnector != null)
            {
                m_ReducerConnector.ViewportDraw(draw);
            }
            if (m_Reducer != null)
            {
                m_Reducer.ViewportDraw(draw);
            }
            if (m_ElbowConnector != null)
            {
                m_ElbowConnector.ViewportDraw(draw);
            }
            if (m_Elbow != null)
            {
                m_Elbow.ViewportDraw(draw);
            }
            if (m_PipeConnector != null)
            {
                m_PipeConnector.ViewportDraw(draw);
            }
            if (m_Pipe != null)
            {
                m_Pipe.ViewportDraw(draw);
            }
        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            if (m_ChangedSetting != UISettings.SettingType.PipingLowerBound)
            {
                // No modeless return code here. Return Cancel
                //
                return SamplerStatus.Cancel;
            }

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

                    // Take new point and update
                    //
                    m_CurrentPos = res.Value;
                    Update();
                    break;

                case PromptStatus.Cancel:
                    return SamplerStatus.Cancel;
            }

            return SamplerStatus.OK;
        }

        public bool Update()
        {
            if (m_LastPair.ObjectId.IsNull || String.IsNullOrEmpty(m_LastPair.Port.Name) && m_Branch == null)
            {
                // We start in space, and so only pipe
                //
                if (m_Pipe != null && m_Pipe.StartPoint.DistanceTo(m_CurrentPos) >= m_Pipe.MinLength)
                {
                    m_Pipe.EndPoint = m_CurrentPos;
                    return true;
                }
                else
                {
                    return false;
                }
            }

            PartSizeProperties lastPart = m_LastPart;
            Autodesk.ProcessPower.PnP3dObjects.Port lastPort = m_LastPair.Port;
            
            if (m_Branch != null)
            {
                // Update branch
                //
                bool bChanged = false;
                if (m_Branch.Update(m_CurrentPos, true, out bChanged))
                {
                    if (bChanged)
                    {
                        // New branch
                        // Recreate all the parts except the branch itself
                        //
                        CreateParts(false);
                    }
                }
                lastPart = m_Branch.BranchProperties; 
                lastPort = m_Branch.BranchEndPort;
            }

            if (m_Reducer != null)
            {
                // Realign
                //
                Matrix3d mat = RoutingHelper.CalculateAttachMatrix(m_ReducerConnector.StartPort,
                                                                   RoutingHelper.CalculateNormal(m_ReducerConnector.StartPort, null),
                                                                   lastPort,
                                                                   RoutingHelper.CalculateNormal(lastPort, null));
                m_ReducerConnector.TransformBy(mat);
                m_Reducer.TransformBy(mat);
                lastPart = m_Reducer.PartSizeProperties; 
                lastPort = m_Reducer.EndPort;
            }

            if (m_Pipe == null)
            {
                // If we have no pipe, nothing to do
                //
                return true;
            }

            // New angle
            //
            Vector3d dir = m_CurrentPos-lastPort.Position;
            double angle = 0.0;
            if (lastPort.Position != m_CurrentPos)
            {
                angle = lastPort.Direction.GetAngleTo(dir);
                System.Diagnostics.Debug.Assert(angle >= 0.0); 
            }

            if (angle > m_ZeroAngle)
            {
                // Do we need an elbow ?
                //
                if (m_bToleranceRouting)
                {
                    // Try connecting pipe directly and use tolerance
                    //
                    if (m_PipeConnector != null && m_Elbow == null)
                    {
                        // Connector already exists and connected to the right part
                        // Only align it with the last port
                        //
                        AttachConnector(m_PipeConnector, lastPort); 
                    }
                    else
                    {
                        // Create new
                        //
                        CreatePipeConnector(lastPart, lastPort);
                    }

                    // Check that connector can satisfy the angle
                    //
                    if (m_PipeConnector != null)
                    {
                        if (angle > m_PipeConnector.SlopeTolerance)
                        {
                            // No, we would need elbow
                            //
                            m_PipeConnector = null;
                        }
                        else
                        {
                            // No elbow
                            //
                            m_CurrentElbow = null;
                            m_Elbow = null;
                            m_ElbowConnector = null;

                            // Override last port direction
                            //
                            lastPort = m_PipeConnector.EndPort;
                            lastPort.Direction = dir;
                            m_PipeConnector.SetOverrideLastPort(lastPort.Position, lastPort.Direction);

                            // Align pipe
                            //
                            AlignPipe(lastPort);
                            return true;
                        }
                    }
                }

                // Try placeing an elbow
                // Can't be too small
                //
                if (angle < m_MinElbowAngle)
                {
                    // No elbow
                    //
                    ClearElbow();
                }
                else
                if (m_bBentPipe)
                {
                    m_CurrentElbow = null;

                    // Is current elbow a bend?
                    //
                    BentPipe bend = m_Elbow as BentPipe;
                    if (bend == null)
                    {
                        // Create new
                        //
                        m_ElbowConnector = null;
                        m_PipeConnector = null;
                        m_Elbow = m_CurrentPipe.CreateBend(angle);
                    }
                    else
                    {
                        // Set angle
                        //
                        bend.Angle = angle;
                    }
                }
                else
                {
                    ElbowPart iElbow = null;
                    if (m_bCutbackElbow)
                    {
                        // Find elbow with an angle greater than required
                        //
                        if (m_ElbowParts != null)
                        {
                            double diff = 0.0;
                            foreach(ElbowPart elbow in m_ElbowParts)
                            {
                                if (elbow.CanCutback && elbow.Angle >= angle)
                                {
                                    double d = elbow.Angle - angle;
                                    if (iElbow == null || d < diff)
                                    {
                                        iElbow = elbow;
                                        diff = d;
                                    } 
                                }  
                            }
                        }
                        if (iElbow != null)
                        {
                            // Do we have cutback elbow already ?
                            //
                            CutbackElbow cutback = m_Elbow as CutbackElbow;

                            // Is it the same elbow ?
                            //
                            if (iElbow == m_CurrentElbow && cutback != null)
                            {
                                // Set angle
                                //
                                cutback.Angle = angle;
                            }
                            else
                            {
                                // Create new
                                //
                                m_CurrentElbow = iElbow;
                                m_ElbowConnector = null;
                                m_PipeConnector = null;
                                m_Elbow = m_CurrentElbow.CreateCutbackElbow(angle);
                            }
                        } 
                    }

                    // If no bend nor cutback, place real elbow
                    //
                    if (iElbow == null)
                    {
                        // Real elbow
                        // Find elbow with an angle less than required
                        //
                        if (m_ElbowParts != null)
                        {
                            double diff = 0.0;
                            foreach(ElbowPart elbow in m_ElbowParts)
                            {
                                if (elbow.Angle <= angle)
                                {
                                    double d = angle - elbow.Angle;
                                    if (iElbow == null || d < diff)
                                    {
                                        iElbow = elbow;
                                        diff = d;
                                    } 
                                }  
                            }
                        }
                        if (iElbow == null)
                        {
                            // No elbow
                            //
                            ClearElbow();
                        }
                        else
                        {
                            // Is it the same elbow ?
                            // If not cutback use it
                            //
                            CutbackElbow cutback = m_Elbow as CutbackElbow;
                            if (iElbow == m_CurrentElbow && m_Elbow != cutback)
                            {
                                // Use
                                //
                            }
                            else
                            {
                                // Creat new
                                //
                                m_CurrentElbow = iElbow;
                                m_ElbowConnector = null;
                                m_PipeConnector = null;
                                m_Elbow = m_CurrentElbow.CreateElbow();
                            }
                        } 
                    }
                }

                if (m_Elbow != null)
                {
                    // Elbow connector
                    //
                    if (m_ElbowConnector != null)
                    {
                        // Connector already exists and connected to the right part
                        // Only align it with the last port
                        //
                        AttachConnector(m_ElbowConnector, lastPort); 
                    }
                    else
                    {
                        // Create new
                        //
                        CreateElbowConnector(lastPart, lastPort);
                    }

                    // Last port
                    //
                    if (m_ElbowConnector != null)
                    {
                        lastPort = m_ElbowConnector.EndPort;
                    }

                    // Align elbow
                    //
                    AlignElbow(lastPort);

                    // Last part/port
                    //
                    lastPart = m_Elbow.PartSizeProperties;
                    Autodesk.ProcessPower.PnP3dObjects.PortCollection ports = m_Elbow.GetPorts(PortType.Static);
                    if (ports.Count > 1)
                    {
                        lastPort = ports[1];
                    }
                }
            }
            else
            {
                // No elbow
                //
                ClearElbow();
            }

            // Pipe connector, if we don't continue previous pipe
            //
            if (!m_bCanContinuePipe || m_Elbow != null)
            {
                m_Pipe.UseStartLength = false;
                if (m_PipeConnector != null)
                {
                    AttachConnector(m_PipeConnector, lastPort); 
                }
                else
                {
                    // Create new
                    //
                    CreatePipeConnector(lastPart, lastPort);
                }
                if (m_PipeConnector != null)
                {
                    lastPort = m_PipeConnector.EndPort;
                }
            }
            else
            {
                // Continue pipe
                //
                m_Pipe.UseStartLength = true;
                m_PipeConnector = null;

                // For asymmetric pipe check that we need to reverse it
                //
                if (m_Pipe.Asymmetric)
                {
                    Autodesk.ProcessPower.PnP3dObjects.PortCollection ports = m_Pipe.GetPorts(PortType.Static);
                    if (!ConnectorWrapper.HasConnection(lastPart, lastPort, m_Pipe.PartSizeProperties, ports[0]))
                    {
                        m_Pipe.Reverse();
                    }
                }
            }

            // Align pipe
            //
            AlignPipe(lastPort);

            return true;
        }

        void AttachConnector (ConnectorWrapper connector, Autodesk.ProcessPower.PnP3dObjects.Port port)
        {
            Autodesk.ProcessPower.PnP3dObjects.Port connectorPort = connector.StartPort;
            Matrix3d mat = RoutingHelper.CalculateAttachMatrix(connectorPort,
                                                               RoutingHelper.CalculateNormal(connectorPort, null),
                                                               port,
                                                               RoutingHelper.CalculateNormal(port, null));
            connector.TransformBy(mat);
        }

        bool CreatePipeConnector (PartSizeProperties part, Autodesk.ProcessPower.PnP3dObjects.Port port)
        {
            if (m_Pipe != null)
            {
                try
                {
                    Autodesk.ProcessPower.PnP3dObjects.PortCollection ports = m_Pipe.GetPorts(PortType.Static);
                    if (m_Pipe.Asymmetric)
                    {
                        // We may need to reverse it
                        //
                        if (!ConnectorWrapper.HasConnection(part, port, m_Pipe.PartSizeProperties, ports[0]))
                        {
                            m_Pipe.Reverse();
                            ports = m_Pipe.GetPorts(PortType.Static);
                        }
                    }

                    // Connect
                    //
                    m_PipeConnector = new ConnectorWrapper();
                    m_PipeConnector.Connect(part, port, m_Pipe.PartSizeProperties, ports[0], null);

                    // Align geometrically
                    //
                    AttachConnector(m_PipeConnector, port); 
                    return true;
                }
                catch (System.Exception ex)
                {
                    if (ex.Message != null)
                    {
                        AcadApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage(ex.Message);
                    }
                }
            }

            m_PipeConnector = null;
            return false;
        }

        void ClearElbow()
        {
            if (m_Elbow != null)
            {
                // We would need to reconnect the pipe without elbow
                //
                m_PipeConnector = null;
            }
            m_CurrentElbow = null;
            m_Elbow = null;
            m_ElbowConnector = null;
        }

        bool CreateElbowConnector (PartSizeProperties part, Autodesk.ProcessPower.PnP3dObjects.Port port)
        {
            if (m_Elbow != null)
            {
                try
                {
                    Autodesk.ProcessPower.PnP3dObjects.PortCollection ports = m_Elbow.GetPorts(PortType.Static);
                    if (ports.Count > 1)
                    {
                        if (m_Elbow.Asymmetric)
                        {
                            // We may need to reverse it
                            //
                            if (!ConnectorWrapper.HasConnection(part, port, m_Elbow.PartSizeProperties, ports[0]))
                            {
                                m_Elbow.Reverse();
                                ports = m_Elbow.GetPorts(PortType.Static);
                            }
                        }

                        // Connect
                        //
                        m_ElbowConnector = new ConnectorWrapper();
                        m_ElbowConnector.Connect(part, port, m_Elbow.PartSizeProperties, ports[0], null);

                        // Align geometrically
                        //
                        AttachConnector(m_ElbowConnector, port);
                        return true;
                    }
                }
                catch (System.Exception ex)
                {
                    if (ex.Message != null)
                    {
                        AcadApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage(ex.Message);
                    }
                }
            }

            m_ElbowConnector = null;
            return false;
        }

        bool AlignElbow(Autodesk.ProcessPower.PnP3dObjects.Port port)
        {
            // Elbow ports
            //
            Autodesk.ProcessPower.PnP3dObjects.PortCollection ports = m_Elbow.GetPorts(PortType.Static);
            if (ports.Count < 2)
            {
                return false;
            }

            // Reference port in the current point
            //
            Autodesk.ProcessPower.PnP3dObjects.Port refPort = new Autodesk.ProcessPower.PnP3dObjects.Port();
            refPort.Position = m_CurrentPos;
            if (port.Position == m_CurrentPos)
            {
                refPort.Direction = -port.Direction;
            }
            else
            {
                refPort.Direction = port.Position - m_CurrentPos;
            }

            // Transform
            //
            Matrix3d mat = RoutingHelper.CalculateAlignMatrix(ports[0], ports[1],
                                                              port, refPort);
            m_Elbow.TransformBy(mat);

            if (m_bCanCutbackPipe)
            {
                // We need to place elbow corner in the last port
                // and so cut back that pipe on ElbowSize + ConnectorSize
                //
                try
                {
                    // Reread ports
                    //
                    ports = m_Elbow.GetPorts(PortType.Static);

                    // Elbow size
                    //
                    double size = m_Elbow.ElbowSize;

                    // Add connector size
                    //
                    if (m_ElbowConnector != null)
                    {
                        size += m_ElbowConnector.StartPort.Position.DistanceTo(m_ElbowConnector.EndPort.Position);
                    }

                    if (size > m_MaxCutbackLength)
                    {
                        // Limit with max possible
                        //
                        size = m_MaxCutbackLength;
                    }
                    
                    // Move
                    //
                    mat = Matrix3d.Displacement(size * ports[0].Direction);
                    m_Elbow.TransformBy(mat);
                    if (m_ElbowConnector != null)
                    {
                        m_ElbowConnector.TransformBy(mat);
                    }
                }
                catch (System.Exception)
                {
                }
            }

            return true;
        }

        bool AlignPipe(Autodesk.ProcessPower.PnP3dObjects.Port port)
        {
            Line3d line = new Line3d(port.Position, port.Direction);
            Point3d pt = line.GetClosestPointTo(m_CurrentPos).Point;
            if (pt == port.Position ||
                pt.DistanceTo(port.Position) < m_Pipe.MinLength ||
                port.Direction.IsCodirectionalTo(port.Position-pt))
            {
                // Make short pipe
                //
                pt = port.Position + m_Pipe.MinLength * port.Direction;
            }

            // Set points
            //
            m_Pipe.ClearPoints();
            m_Pipe.StartPoint = port.Position;
            m_Pipe.EndPoint = pt;
            return true;
        }

        bool DoAutoRoute (out bool bCancelled)
        {
            bCancelled = false;
            Editor ed                   = AcadApp.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                using (AutoRoute ar = new AutoRoute (m_LastPair, m_SnapPair,
                                              m_Spec, m_Size, m_CurrentPipe, m_ElbowParts,
                                              m_bStubin, m_bCutbackElbow, m_bBentPipe))
                {
                    if (ar.PathCount == 0)
                    {
                        return false;
                    }
                    else
                    if (ar.PathCount == 1)
                    {
                        // Single path. Just accept it
                        //
                        ar.CurrentPath = 0;
                    }
                    else
                    {
                        // Select path with preview
                        //
                        PromptKeywordOptions opts = new PromptKeywordOptions("");
                        opts.Keywords.Add("Accept");
                        opts.Keywords.Add("Next");
                        opts.Keywords.Add("Previous");
                        opts.Keywords.Add("Undo");
                        opts.Keywords.Default = "Accept";
                        int idx = 0;
                        while (true)
                        {
                            opts.Message = "\nConnect or preview next solution " + (idx+1).ToString() + " of " + ar.PathCount.ToString();
                            ar.CurrentPath = idx;
                            ar.Preview();

                            PromptResult pRes = ed.GetKeywords(opts);
                            if (pRes.Status != PromptStatus.OK || pRes.StringResult == "Undo")
                            {
                                bCancelled = true;
                                break;
                            }
                            else
                            if (pRes.StringResult == "Accept")
                            {
                                break;
                            }
                            else
                            if (pRes.StringResult == "Next")
                            {
                                idx++;
                                if (idx == ar.PathCount)
                                {
                                    idx = 0;
                                }
                                opts.Keywords.Default = "Next";
                            }
                            else
                            if (pRes.StringResult == "Previous")
                            {
                                idx--;
                                if (idx < 0 )
                                {
                                    idx = ar.PathCount-1;
                                }
                                opts.Keywords.Default = "Previous";
                            }
                        }

                        ar.EndPreview();
                    }

                    if (bCancelled)
                    {
                        return false;
                    }
                    else
                    {
                        // Append
                        //
                        if (m_GroupId < 0)
                        {
                            m_GroupId = PnP3dTagFormat.findOrCreateNewLineGroup(m_LineNumber);
                        }
                        ar.Append(m_GroupId);
                        return true;
                    }
                }
            }
            catch (System.Exception ex)
            {
                // No paths
                //
                ed.WriteMessage("\nCannot create branch");
                if (ex.Message != null)
                {
                    ed.WriteMessage(ex.Message);
                }
                return false;
            }
        }
    }

    public class PlaceFitting
    {
        public void Run()
        {
            Editor ed       = AcadApp.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                Database db                 = AcadApp.DocumentManager.MdiActiveDocument.Database;
                Project currentProject      = PlantApp.CurrentProject.ProjectParts["Piping"];
                DataLinksManager dlm        = currentProject.DataLinksManager;
                DataLinksManager3d dlm3d    = DataLinksManager3d.Get3dManager(dlm);

                // Select a fitting in the current spec with the current size
                //
                UISettings sett = new UISettings();
                String lineNumber = sett.CurrentLineNumber;
                NominalDiameter nd = NominalDiameter.FromDisplayString(null, sett.CurrentSize);
                String spec = sett.CurrentSpec;
                SpecManager specMgr = SpecManager.GetSpecManager();
                String[] partTypes = specMgr.GetAvailablePartTypes(spec);
                if (partTypes == null)
                {
                    return;
                }

                // Ask for a part type
                //
                List<String> types = new List<String>();
                List<String> kwds = new List<String>();
                int cnt = 0;
                foreach (String type in partTypes)
                {
                    // Skip pipes and valve actuators
                    //
                    if (String.Compare(type, "Pipe", true) == 0 ||
                        String.Compare(type, "ValveActuator", true) == 0)
                    {
                        continue;
                    }

                    // Also, skip fasteners
                    //
                    if (!dlm.HasTable(type))
                    {
                        continue;
                    }
                    PnPTable tbl = dlm.GetPnPDatabase().Tables[type];
                    if (tbl.IsKindOf("Fasteners"))
                    {
                        continue;
                    }

                    // Make a keyword in the form "n.Type"
                    // And kill spaces and other special symbols
                    //
                    cnt += 1;
                    String s = type.Replace(" ", "");
                    s = s.Replace("/", "");
                    s = s.Replace(".", "");
                    kwds.Add(cnt.ToString() + "." + s);
                    types.Add(type);
                }
                PromptResult res = ed.GetKeywords("\nSelect part type", kwds.ToArray());
                if (res.Status != PromptStatus.OK)
                {
                    return;
                }
                String partType = types[kwds.IndexOf(res.StringResult)];
                
                // Parts
                //
                List<PartSizeProperties> parts = new List<PartSizeProperties>();
                SpecPartReader reader = specMgr.SelectParts(spec, partType, nd);
                while (reader.Next())
                {
                    SpecPart part = new SpecPart();
                    part.AssignFrom(reader.Current);
                    parts.Add(part);
                }

                // Ask for a part
                //
                kwds.Clear();
                cnt = 0;
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
                            s = s.Replace(",", "");
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
                res = ed.GetKeywords("\nSelect part", kwds.ToArray());
                if (res.Status != PromptStatus.OK)
                {
                    return;
                }
                PartSizeProperties fittingPart = parts[kwds.IndexOf(res.StringResult)];

                // Create entity
                //
                ContentManager conMgr = ContentManager.GetContentManager();
                Units lUnit = (Units)Autodesk.ProcessPower.AcPp3dObjectsUtils.ProjectUnits.CurrentLinearUnit;
                ObjectId symbolId = conMgr.GetSymbol(fittingPart, db, lUnit);
                PipeInlineAsset fitting = new PipeInlineAsset();
                fitting.SymbolId = symbolId;

                // Place
                //
                while(true)
                {
                    // Drag. Ask for a placement point
                    //
                    FittingDragPoint dragPoint = new FittingDragPoint(fitting, fittingPart);
                    dragPoint.StartPointMonitor();
                    PromptStatus st = ed.Drag(dragPoint).Status;
                    dragPoint.StopPointMonitor();
                    if (st != PromptStatus.OK)
                    {
                        return;
                    }
                    fitting = dragPoint.Fitting;
                    fittingPart = dragPoint.Props;

                    if (dragPoint.SnapId.IsNull)
                    {
                        // No snapping. Ask for a rotation
                        //
                        FittingDragAngle dragAngle = new FittingDragAngle(fitting);
                        ed.Drag(dragAngle);
                    }

                    try
                    {
                        // Line group
                        //
                        int groupId = -1;
                        if (dragPoint.SnapId.IsNull)
                        {
                            // Current
                            //
                            groupId = PnP3dTagFormat.findOrCreateNewLineGroup(lineNumber);
                        }
                        else
                        {
                            // From snap entity
                            //
                            groupId = PnP3dTagFormat.lineGroupIdFromObjId(dragPoint.SnapId);
                        }
                        
                        // If we don't need a branch,  create fitting
                        //
                        if (!dragPoint.IsOlet || dragPoint.SnapId.IsNull)
                        {
                            // Make a copy to place
                            //
                            PipeInlineAsset fittingCopy = (PipeInlineAsset)fitting.Clone();

                            // Add
                            //
                            ObjectId id;
                            using (Transaction tr = db.TransactionManager.StartOpenCloseTransaction())
                            {
                                using (PipingObjectAdder pipeObjAdder = new PipingObjectAdder(dlm3d, db))
                                {
                                    pipeObjAdder.LineGroupId = groupId;
                                    pipeObjAdder.Add(fittingPart, fittingCopy);
                                    tr.AddNewlyCreatedDBObject(fittingCopy, true);
                                    id = fittingCopy.ObjectId;
                                }

                                tr.Commit();
                            }

                            if (!dragPoint.SnapId.IsNull)
                            {
                                // Connect to the part
                                //
                                if (!String.IsNullOrEmpty(dragPoint.SnapPort.Name))
                                {
                                    // Connect ports
                                    //
                                    Pair pair1 = new Pair();
                                    pair1.ObjectId = id;
                                    pair1.Port = dragPoint.Ports[dragPoint.Port1];
                                    Pair pair2 = new Pair();
                                    pair2.ObjectId = dragPoint.SnapId;
                                    pair2.Port = dragPoint.SnapPort;

                                    // Line group from that osnap ent
                                    //

                                    // Lock snapped entity, so it won't move
                                    //
                                    LockEntity eLock = new LockEntity(dragPoint.SnapId, true, true, true);

                                    RoutingHelper.Connect(pair1, pair2, false, false, false, groupId, null, Tolerance.Global);
                                }
                                else
                                {
                                    // Break pipe
                                    //
                                    RoutingHelper.BreakPipeWithAsset(dragPoint.SnapId, id, dragPoint.Port1,  dragPoint.Port2, true);
                                }
                            }

                            // TODO: rotate fitting
                        }
                        else
                        {
                            // Create a branch
                            //
                            BranchWrapper branch = new BranchWrapper(dragPoint.SnapId, dragPoint.SnapPort, fittingPart, fitting.SymbolId);
                            if (dragPoint.Port2 >= 0)
                            {
                                // Orient
                                //
                                bool bBranchChanged;
                                branch.Update(dragPoint.Ports[dragPoint.Port2].Position, false, out bBranchChanged);
                            }

                            // Append
                            //
                            branch.Append(groupId);

                            // TODO: rotate branch
                        }
                    }
                    catch (System.Exception ex)
                    {
                        if (ex.Message != null)
                        {
                            ed.WriteMessage(ex.Message);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                if (ex.Message != null)
                {
                    ed.WriteMessage(ex.Message);
                }
            }
        }
    }

    // Dragger classes
    //
    public class FittingDragPoint : EntityJig
    {
        Point3d m_CurrentPos;
        bool m_bNewSnap;
        JigPromptPointOptions m_Opts;

        public PipeInlineAsset Fitting
        {
            get
            {
                return (PipeInlineAsset)Entity;
            }
        }

        public PartSizeProperties Props
        {
            get;
            set;
        }

        public Autodesk.ProcessPower.PnP3dObjects.PortCollection Ports
        {
            get
            {
                return ((PipeInlineAsset)Entity).GetPorts(PortType.Static);
            }
        }

        public int Port1
        {
            get;
            set;
        }

        public int Port2
        {
            get;
            set;
        }

        NominalDiameter ND
        {
            get
            {
                if (Port1 >= 0)
                {
                    // Port ND
                    //
                    Autodesk.ProcessPower.PnP3dObjects.SpecPort port = Props.Port(Ports[Port1].Name);
                    return port.NominalDiameter;
                }

                // Part ND
                //
                return Props.NominalDiameter;
            }
        }

        public bool IsOlet
        {
            get;
            set;
        }

        public bool IsElbolet
        {
            get;
            set;
        }

        public bool IsLatrolet
        {
            get;
            set;
        }


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
        
        public PartSizeProperties SnapProps
        {
            get;
            set;
        }

        NominalDiameter SnapND
        {
            get
            {
                if (!String.IsNullOrEmpty(SnapPort.Name))
                {
                    // Port ND
                    //
                    Autodesk.ProcessPower.PnP3dObjects.SpecPort port = SnapProps.Port(SnapPort.Name);
                    return port.NominalDiameter;
                }

                // Part ND
                //
                return SnapProps.NominalDiameter;
            }
        }

        public FittingDragPoint(PipeInlineAsset fitting, PartSizeProperties part)
            : base(fitting)
        {
            Props = part;

            // Olet ?
            //
            IsOlet = false;
            IsElbolet = false;
            IsLatrolet = false;
            if (String.Compare(part.Type, "Olet", true) == 0)
            {
                IsOlet = true;

                // Special cases: elbolet and latrolet
                //
                String subType = part.PropValue("PartSubType") as String;
                if (String.Compare(subType, "Elbolet", true) == 0)
                {
                    IsElbolet = true;
                }
                else
                if (String.Compare(subType, "Latrolet", true) == 0)
                {
                    IsLatrolet = true;
                }
            }

            // Fitting ports we will be using for snap
            // TODO: tee, cross... Port2 could be not the second one
            //
            switch (Ports.Count)
            {
            case 0:
                Port1 = -1;
                Port2 = -1;
                break;
            case 1:
                Port1 = 0;
                Port2 = -1;
                break;
            default:
                Port1 = 0;
                Port2 = 1;
                break;
            }

            m_CurrentPos = fitting.Position;
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
                    // Resize ?
                    //
                    if (!ND.Equals(SnapND))
                    {
                        // Yes, select a part from the same family with new nd
                        //
                        bool bFound = false;
                        String partFamilyId = Props.PropValue("PartFamilyId") as String;
                        SpecManager specMgr = SpecManager.GetSpecManager();
                        SpecPartReader reader = specMgr.SelectParts(Props.Spec, Props.Type, SnapND);
                        while (reader.Next())
                        {
                            String newId = reader.Current.PropValue("PartFamilyId") as String;
                            if (String.Compare(partFamilyId, newId, true) == 0)
                            {
                                // Create new block
                                //
                                ContentManager conMgr = ContentManager.GetContentManager();
                                Database db = AcadApp.DocumentManager.MdiActiveDocument.Database;
                                ObjectId symbolId = conMgr.GetSymbol(reader.Current, db, Props.LinearUnit);
                                if (!symbolId.IsNull)
                                {
                                    // Success
                                    //
                                    bFound = true;
                                    Fitting.SymbolId = symbolId;
                                    SpecPart part = new SpecPart();
                                    part.AssignFrom(reader.Current);
                                    Props = part;
                                }
                            }
                        }

                        if (!bFound)
                        {
                            // Clear snap
                            //
                            SnapId = ObjectId.Null;
                        }
                    }
                }
            }

            // Still have osnap after resizing?
            //
            if (!SnapId.IsNull)
            {
                // Align
                //
                if (!IsOlet)
                {
                    // Snap to port, which is either real port or point on a pipe with pipe dir
                    //
                    Autodesk.ProcessPower.PnP3dObjects.Port fittingPort = Ports[Port1];
                    Matrix3d mat = RoutingHelper.CalculateAttachMatrix(fittingPort,
                                                                       RoutingHelper.CalculateNormal(fittingPort, null),
                                                                       SnapPort,
                                                                       RoutingHelper.CalculateNormal(SnapPort, null));
                    Fitting.TransformBy(mat);
                }
                else
                {
                    // TODO:
                    //
                    Fitting.Position = m_CurrentPos;
                }
            }
            else
            {
                // No osnap. Just set new point
                //
                Fitting.Position = m_CurrentPos;
            }

            return true;
        }

        public void StartPointMonitor()
        {
            SnapId = ObjectId.Null;
            AcadApp.DocumentManager.MdiActiveDocument.Editor.PointMonitor += new PointMonitorEventHandler(drag_PointMonitorEventHandler);
        }

        public void StopPointMonitor()
        {
            AcadApp.DocumentManager.MdiActiveDocument.Editor.PointMonitor -= new PointMonitorEventHandler(drag_PointMonitorEventHandler);
        }

        public void drag_PointMonitorEventHandler(object sender, PointMonitorEventArgs args)
        {
            if (Port1 < 0)
            {
                // Fitting has no ports. Nothing to snap
                //
                return;
            }

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

            Database db                 = AcadApp.DocumentManager.MdiActiveDocument.Database;              
            ConnectionManager conMan    = new ConnectionManager();

            foreach (FullSubentityPath path in paths)
            {
                ObjectId []ids = path.GetObjectIds();
                if (ids.Length == 0)
                {
                    continue;
                }
                ObjectId id = ids[0];

                // Can snap to pipe, fitting and equipment
                //
                if (!id.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Pipe))) &&
                    !id.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(PipeInlineAsset))) &&
                    !id.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Equipment))))
                {
                    continue;
                }

                // Get ports and properties
                //
                Autodesk.ProcessPower.PnP3dObjects.PortCollection ports = null;
                bool bEditable = false;
                Vector3d dir = Vector3d.XAxis;
                Transaction tr = null;
                try
                {
                    tr = db.TransactionManager.StartOpenCloseTransaction();
                    Autodesk.ProcessPower.PnP3dObjects.Part part = tr.GetObject(id, OpenMode.ForRead) as Autodesk.ProcessPower.PnP3dObjects.Part;
                    ports = part.GetPorts(PortType.Static);
                    SnapProps = part.PartSizeProperties;

                    // And check that the part is editable (not from xref and can be open for write) 
                    //
                    if (part.OwnerId == db.CurrentSpaceId)
                    {
                        tr.GetObject(id, OpenMode.ForWrite);
                        bEditable = true;
                    }

                    // And for the pipe, get its dir
                    //
                    Pipe pipe = part as Pipe;
                    if (pipe != null)
                    {
                        dir = pipe.EndPoint - pipe.StartPoint;
                    }
                }
                catch (System.Exception)
                {
                }
                finally
                {
                    if (tr != null)
                    {
                        tr.Commit();
                    }
                }

                if (!IsOlet)
                {
                    // Can snap to open port
                    //
                    if (ports != null)
                    {
                        foreach (Autodesk.ProcessPower.PnP3dObjects.Port port in ports)
                        {
                            if (port.Position != snapPoint)
                            {
                                continue;
                            }

                            Pair pair = new Pair();
                            pair.ObjectId = id;
                            pair.Port = port;

                            // Connected ?
                            //
                            if (!conMan.IsConnected(pair))
                            {
                                // Take this port
                                //
                                SnapId = id;
                                SnapPort = port;
                                if (id != oldSnapId)
                                {
                                    m_bNewSnap = true;
                                }
                                return;
                            }
                        }
                    }
                }

                // Check branching pipes and fittings
                //
                if (id.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Equipment))))
                {
                    continue;
                }
                if (!bEditable)
                {
                    continue;
                }

                // Make port
                //
                SnapPort = new Autodesk.ProcessPower.PnP3dObjects.Port();
                SnapPort.Position = snapPoint;
                SnapPort.Direction = dir;

                if (!IsOlet)
                {
                    if (Port2 >= 0)
                    {
                        // Can break pipe with inline asset
                        //
                        if (id.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Pipe))))
                        {
                            SnapId = id;
                        }
                    }
                }
                else
                if (IsLatrolet)
                {
                    // Pipes and fittings except elbows
                    //
                    if (String.Compare(SnapProps.Type, "Elbow", true) != 0 &&
                        String.Compare(SnapProps.Type, "PipeBend", true) != 0)
                    {
                        SnapId = id;
                    }
                }
                else
                if (IsElbolet)
                {
                    // Only elbows
                    //
                    if (String.Compare(SnapProps.Type, "Elbow", true) == 0 ||
                        String.Compare(SnapProps.Type, "PipeBend", true) == 0)
                    {
                        SnapId = id;
                    }
                }
                else
                {
                    // Fitting or pipe
                    //
                    SnapId = id;
                }

                if (!SnapId.IsNull)
                {
                    // We have snap
                    //
                    if (id != oldSnapId)
                    {
                        m_bNewSnap = true;
                    }
                    break;
                }
            }
        }
    }

    public class FittingDragAngle : EntityJig
    {
        Double m_CurrentAngle;
        Double m_DefaultAngle;
        Point3d m_Position;
        Vector3d m_ZAxis;
        Matrix3d m_Trans;
        JigPromptAngleOptions m_Opts;

        public FittingDragAngle(PipeInlineAsset fitting)
            : base(fitting)
        {
            m_CurrentAngle = m_DefaultAngle = Vector3d.XAxis.GetAngleTo(fitting.XAxis, fitting.ZAxis);
            m_Position = fitting.Position;
            m_ZAxis = fitting.ZAxis;
            m_Trans = Matrix3d.Identity;

            m_Opts = new JigPromptAngleOptions("\nSelect rotation angle");
            m_Opts.UserInputControls = UserInputControls.NullResponseAccepted;
            m_Opts.Cursor |= CursorType.RubberBand;
            m_Opts.BasePoint = fitting.Position;
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

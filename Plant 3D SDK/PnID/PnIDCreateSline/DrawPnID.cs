////
//// (C) Copyright 2013 by Autodesk, Inc.
////
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
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Reflection;

using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;

using Autodesk.ProcessPower.PnIDObjects;
using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.ProjectManager;
using Autodesk.ProcessPower.PnIDGUIUtil;
using Autodesk.ProcessPower.Styles;

namespace PnIDCreateSline
{
    public class CreatePnID
    {
        public static void CreateSline_Automatic()
        {
            Initialize();            

            // Create a start asset: Lets say a Tank
            //
            AddAsset(new Point3d(7.25, 9, 0), Vector3d.XAxis, Vector3d.ZAxis, new Scale3d(1), "Dome Roof Tank Style", String.Empty, true);

            // Create a start asset: Lets say a Pump
            //
            AddAsset(new Point3d(20.3125, 15, 0), Vector3d.XAxis, Vector3d.ZAxis, "Horizontal Centrifugal Pump Style", String.Empty, true);

            // Create a sline between the two assets: Primary PipeLine
            // Note that the flow-arrows will be automatically created (Based on the setting on the sline style)
            // Note that intersection gaps will also be automatically created.
            // During sline creation, all connections are made. Nozzles are created if required.
            //
            Point3dCollection points = new Point3dCollection();
            points.Add(new Point3d(10, 10, 0));
            points.Add(new Point3d(15, 10, 0));
            points.Add(new Point3d(15, 15, 0));
            points.Add(new Point3d(20, 15, 0));
            AddSline(points, "New Primary Style", true);

            // Create an inline asset: Say a gate valve
            //
            AddAsset(new Point3d(12, 10, 0), Vector3d.XAxis, Vector3d.ZAxis, "Gate Valve Style", String.Empty, true);

            // Create another inline asset: Say a control valve
            //
            AddAsset(new Point3d(15, 12.5, 0), Vector3d.XAxis, Vector3d.ZAxis, "Globe Valve Style", "Hand Wheel With Actuator Style", true);
            
            Terminate();
        }

        public static void CreateSline_Manual()
        {
            Initialize();

            // Create a start asset: Lets say a Tank
            //
            int iTankId = AddAsset(new Point3d(7.25, 9, 0), Vector3d.XAxis, Vector3d.ZAxis, new Scale3d(1), "Dome Roof Tank Style", String.Empty, false);

            // Create a flanged nozzle and attach to the tank
            //
            int iFlangedNozzleId = AddNozzle(iTankId, "Flanged Nozzle Style", new Point3d(10, 10, 0), Vector3d.XAxis, Vector3d.ZAxis, false);

            // Create a start asset: Lets say a Pump
            //
            int iPumpId = AddAsset(new Point3d(20.3125, 15, 0), Vector3d.XAxis, Vector3d.ZAxis, "Horizontal Centrifugal Pump Style", String.Empty, false);

            // Create a flanged nozzle and attach to the tank
            //
            int iAssumedNozzleId = AddNozzle(iPumpId, "Assumed Nozzle Style", new Point3d(20, 15, 0), Vector3d.XAxis, Vector3d.ZAxis, false);

            // Create a sline between the two assets: Primary PipeLine
            // Note that the flow-arrows will be automatically created (Based on the setting on the sline style)
            // Note that intersection gaps will also be automatically created.
            // During sline creation, all connections are made. Nozzles are created if required.
            //
            Point3dCollection points = new Point3dCollection();
            points.Add(new Point3d(10.1875, 10, 0));
            points.Add(new Point3d(15, 10, 0));
            points.Add(new Point3d(15, 15, 0));
            points.Add(new Point3d(20, 15, 0));
            int iLineId = AddSline(points, "New Primary Style", false);

            // Connect endline asset (tank) to line via nozzle
            //
            ConnectLineWithAsset(iLineId, iFlangedNozzleId, AssetContext.ContextStart);

            // Connect endline asset (pump) to line via nozzle
            //
            ConnectLineWithAsset(iLineId, iAssumedNozzleId, AssetContext.ContextEnd);

            // Create an inline asset: Say a gate valve
            //
            int iGateValveAssetId = AddAsset(new Point3d(12, 10, 0), Vector3d.XAxis, Vector3d.ZAxis, "Gate Valve Style", String.Empty, false);

            // Connect inline asset to line
            //
            ConnectLineWithAsset(iLineId, iGateValveAssetId, AssetContext.ContextInline);

            // Create another inline asset: Say a control valve
            //
            int iControlValveAssetId = AddAsset(new Point3d(15, 12.5, 0), Vector3d.YAxis, Vector3d.ZAxis, "Globe Valve Style", "Hand Wheel With Actuator Style", false);

            // Connect inline asset to line
            //
            ConnectLineWithAsset(iLineId, iControlValveAssetId, AssetContext.ContextInline);

            Terminate();
        }

        public static bool Initialize()
        {
            if (PlantApplication.CurrentProject == null)
                return false;

            Helper.PnIdProject = (PnIdProject)PlantApplication.CurrentProject.ProjectParts["PnId"];
            Helper.ActiveDataLinksManager = Helper.PnIdProject.DataLinksManager;
            Helper.ActiveDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            return true;
        }

        public static void Terminate()
        {
            Helper.PnIdProject = null;
            Helper.ActiveDataLinksManager = null;
            Helper.ActiveDocument = null;
        }

        public static int AddSline(Point3dCollection points, String sStyleName, bool bEnableAutomations)
        {
            int lineID = -1;
            try
            {
                using (LineSegmentAdder lineSegmentAdder = new LineSegmentAdder(Helper.ActiveDocument.Database, bEnableAutomations))
                {
                    // We want to keep each line segment separate in its own group.
                    //
                    lineSegmentAdder.MergeGroups = false;

                    // We also do not want to merge lines that are collinear.
                    //
                    lineSegmentAdder.MergeLines = false;

                    // Create new line segment.
                    //
                    LineSegment lineSegment = new LineSegment();

                    // Get the style id for the given class name and set it on the line segment.
                    //
                    ObjectId StyleId = StyleUtil.GetStyleIdFromName(sStyleName, StyleType.eGraphical);
                    if (StyleId == ObjectId.Null)
                        return lineID;
                    lineSegment.StyleID = StyleId;

                    // Add all vertexes
                    //
                    for (int i = 0; i < points.Count; i++)
                        lineSegment.AddVertexAt(i, points[i]);

                    // Make sure atleast 2 vertices are available
                    //
                    if (lineSegment.Vertices.Count < 2)
                        return lineID;

                    // Add line to db
                    //
                    lineID = lineSegmentAdder.Add(lineSegment, -1);
                }
            }
            catch
            {

            }
            return lineID;
        }

        public static int AddAsset(Point3d insertionPt, Vector3d xAxis, Vector3d Normal, String sPrimaryStyleName, String sSecondaryStyleName, bool bEnableAutomations)
        {
            // Get the style id for the given class name and set it on the line segment.
            //
            double dScaleFactor = 1.0;
            ObjectId styleIdPrimary = StyleUtil.GetStyleIdFromName(sPrimaryStyleName, StyleType.eGraphical);
            if (styleIdPrimary != ObjectId.Null)
                dScaleFactor = GetAssetScale(styleIdPrimary);

            return AddAsset(insertionPt, xAxis, Normal, new Scale3d(dScaleFactor), sPrimaryStyleName, sSecondaryStyleName, bEnableAutomations);
        }

        public static int AddAsset(Point3d insertionPt, Vector3d xAxis, Vector3d Normal, Scale3d scaleFactor, String sPrimaryStyleName, String sSecondaryStyleName, bool bEnableAutomations)
        {
            int iAssetId = -1;
            try
            {
                // Create the inline asset on the line segment.
                //
                using (var assetAdder = new AssetAdder(Helper.ActiveDocument.Database, bEnableAutomations))
                {
                    // Get the style id for the given class name and set it on the line segment.
                    //
                    ObjectId styleIdPrimary = StyleUtil.GetStyleIdFromName(sPrimaryStyleName, StyleType.eGraphical);
                    if (styleIdPrimary == ObjectId.Null)
                        return iAssetId;

                    // Add as control valve if sSecondaryClassName is NOT null or empty.
                    // Only ControlValves will have this value set.
                    //
                    if (!String.IsNullOrEmpty(sSecondaryStyleName))
                    {
                        double rotationAngle = xAxis.GetAngleTo(Vector3d.XAxis);
                        iAssetId = assetAdder.AddControlValve(insertionPt, sPrimaryStyleName, sSecondaryStyleName, rotationAngle/*, scaleFactor*/);
                    }
                    // Add as ordinary Inline asset (NOT a control valve)
                    //
                    else
                    {
                        Asset inlineAsset = null;
                        String sClassName = StyleUtil.GetClassFromStyleType(sPrimaryStyleName, StyleType.eGraphical, true);
                        if (Utils.isKindOfClass(sClassName, /*NOXLATE*/"Connectors"))
                            inlineAsset = new DynamicOffPage();
                        else if (AssetUtil.IsDynamicAsset(styleIdPrimary))
                            inlineAsset = new DynamicAsset();
                        else
                            inlineAsset = new Asset();

                        inlineAsset.SetDatabaseDefaults();
                        inlineAsset.StyleId = styleIdPrimary;
                        inlineAsset.Position = insertionPt;
                        inlineAsset.ScaleFactors = scaleFactor;
                        inlineAsset.XAxis = xAxis;
                        inlineAsset.Normal = Normal;

                        iAssetId = assetAdder.Add(inlineAsset);

                        // Initilialize dynamic asset to get annotative behaviour.
                        //
                        if (AssetUtil.IsDynamicAsset(styleIdPrimary))
                        {
                            InitializeDynamicAsset(iAssetId);
                        }
                    }
                }
            }
            catch
            {

            }
            return iAssetId;
        }

        private static int AddNozzle(int iOwnerId, String sNozzleStyleName, Point3d insertionPt, Vector3d xAxis, Vector3d normal, bool bEnableAutomations)
        {
            int nozzleId = -1;
            using (var assetAdder = new AssetAdder(Helper.ActiveDocument.Database, false))
            {
                try
                {
                    // Add nozzle and connect to equipment
                    // 
                    double rotationAngle = xAxis.GetAngleTo(Vector3d.XAxis);
                    nozzleId = assetAdder.AddNozzleToAsset(iOwnerId, sNozzleStyleName, insertionPt, rotationAngle);
                 }
                catch
                {
                }
            }
            return nozzleId;
        }

        public static double GetAssetScale(ObjectId assetStyleId)
        {
            double dScaleFactor = 1.0;
            using (Transaction modelTrans = Helper.ActiveDocument.Database.TransactionManager.StartTransaction())
            {
                try
                {
                    BlockTable modelBT = (BlockTable)modelTrans.GetObject(Helper.ActiveDocument.Database.BlockTableId, OpenMode.ForRead, false);
                    if (modelBT != null)
                    {
                        BlockTableRecord modelBTR = (BlockTableRecord)modelTrans.GetObject(modelBT[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                        if (modelBTR != null)
                        {
                            using (AssetStyle assetStyle = (AssetStyle)modelTrans.GetObject(assetStyleId, OpenMode.ForRead))
                            {
                                if (assetStyle != null)
                                    dScaleFactor = assetStyle.ScaleFactor;
                            }
                        }
                        modelBTR.Dispose();
                    }                    
                    modelBT.Dispose();
                }
                finally
                {
                    // Nothing to save here.
                    //
                    modelTrans.Abort();
                }
            }
            return dScaleFactor;
        }

        public static void InitializeDynamicAsset(int iAssetId)
        {
            using (Transaction modelTrans = Helper.ActiveDocument.Database.TransactionManager.StartTransaction())
            {
                try
                {
                    BlockTable modelBT = (BlockTable)modelTrans.GetObject(Helper.ActiveDocument.Database.BlockTableId, OpenMode.ForWrite, false);
                    if (modelBT != null)
                    {
                        BlockTableRecord modelBTR = (BlockTableRecord)modelTrans.GetObject(modelBT[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                        if (modelBTR == null)
                            return;

                        ObjectId iAssetObjectId = Helper.ConvertRowIdToObjectId(iAssetId);
                        if (!iAssetObjectId.IsNull)
                        {
                            using (DynamicAsset dynAsset = (DynamicAsset)modelTrans.GetObject(iAssetObjectId, OpenMode.ForWrite))
                            {
                                if (dynAsset != null)
                                    dynAsset.Initialize();
                            }
                        }
                        modelBTR.Dispose();
                    }
                    modelBT.Dispose();
                    modelTrans.Commit();
                }
                catch
                {
                    modelTrans.Abort();
                }
            }
        }

        public static bool ConnectLineWithAsset(int iLineId, int iAssetId, AssetContext currentAssetContext)
        {
            bool bRet = false;
            try
            {
                LineUtil.ConnectLineWithAsset(iLineId, iAssetId, currentAssetContext);
                bRet = true;
            }
            catch
            {
                bRet = false;
            }
            return bRet;
        }
    }
}
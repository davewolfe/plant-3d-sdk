//
// (C) Copyright 2009 by Autodesk, Inc.
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

using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.ProcessPower.Validation;
using Autodesk.ProcessPower.PnIDDwgValidation;
using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.ProjectManager;
using Autodesk.ProcessPower.PnP3dObjects;

namespace Autodesk.ProcessPower.PipingValidation
{
    public class UnconnectedPortRule : DrawingRule, IAcPpDrawingValidationRule
    {
        public UnconnectedPortRule()
            : base(RuleDescription, RuleGuid, RuleName)
        {
        }

        public override bool GetEnabledFromCurrentProject()
        { 
            //return GetEnabledFromPipingPart(this.Guid); 
            return true;
        }

        public override void SetEnabledToCurrentProject(bool bEnabled)
        {
            //SetEnabledToPipingPart(this.Guid, bEnabled); 
        }

        public override void SubEvaluate(Database dwg)
        {
            if (dwg == null)
            {
                return;
            }

            ObjectIdCollection arrAll3dObjIds = PipingValidationUtils.GetAll3dObjectIds(dwg);
            foreach (ObjectId idPart in arrAll3dObjIds)
            {
                using (Transaction transaction = dwg.TransactionManager.StartTransaction())
                {
                    Part part = transaction.GetObject(idPart, OpenMode.ForRead) as Part;
                    if (part == null)
                    {
                        continue;
                    }

                    ConnectionManager cmgr = new ConnectionManager();
                    if (cmgr == null)
                    {
                        continue;
                    }

                    PortCollection ports = part.GetPorts(PortType.All);
                    foreach (Port port in ports)
                    {
                        Pair pair = new Pair();
                        pair.ObjectId = idPart;
                        pair.Port = port;

                        if (!cmgr.IsConnected(pair))
                        {
                            UnconnectedPortError error = new UnconnectedPortError(dwg.FingerprintGuid);
                            error.DisplayName = "Open port [" + port.Name + "] on component";
                            error.Description = "The port is not connected to a component.";

                            string strPartTag = PipingValidationUtils.GetTagValue(idPart);
                            if (!String.IsNullOrEmpty(strPartTag))
                            {
                                error.DisplayName = error.DisplayName + " [" + strPartTag +"]";
                            }

                            error.ObjectId = PipingValidationUtils.GetPpObjectId(idPart);
                            error.Point = port.Position;

                            AcPpValidationManager vmgr = ValidationSingleton.Manager;
                            vmgr.Errors.Add(error);
                        }
                    }
                }
            }
        }
                
        public static string RuleGuid = "Autodesk.ProcessPower.PipingValidation.UnconnectedPortRule";
        public static string RuleName = "UnconnectedPortRule";
        public static string RuleDescription = "Unconnected Port Rule";
    }

    public class UnconnectedPortError : DrawingError
    {
        public UnconnectedPortError(string strDrawingGuid)
            : base(ErrorName, ErrorDiscription, UnconnectedPortRule.RuleGuid, strDrawingGuid)
        {
        }

        public override LinkedList<ValidationDetail> Details
        { 
            get
            {
                LinkedList<ValidationDetail> properties = new LinkedList<ValidationDetail>();
                ValidationDetail currentproperty = new ValidationDetail();

                currentproperty.Field = "Error Type";
                currentproperty.Value = "Open port";
                currentproperty.HelperString = "The port is not connected to a component.";

                properties.AddLast(currentproperty);

                return properties;
            }
        }

        public static string ErrorName = "Open Ports";
        public static string ErrorDiscription = "Unconnected Port Error";
    }

    public class PipingValidationUtils
    {
        public static ObjectIdCollection GetAll3dObjectIds(Database dwg)
        {
            ObjectIdCollection arr3dObjectIds = null;

            using (Transaction transaction = dwg.TransactionManager.StartTransaction())
            {
                BlockTable bt = transaction.GetObject(dwg.BlockTableId, OpenMode.ForRead) as BlockTable;
                if (bt != null)
                {
                    BlockTableRecord btrModelSpace = transaction.GetObject(bt["*Model_Space"], OpenMode.ForRead) as BlockTableRecord;
                    if (btrModelSpace != null)
                    {
                        arr3dObjectIds = new ObjectIdCollection();

                        BlockTableRecordEnumerator btre = btrModelSpace.GetEnumerator();
                        while (btre.MoveNext())
                        {
                            Part part = transaction.GetObject(btre.Current, OpenMode.ForRead) as Part;
                            if (part != null)
                            {
                                arr3dObjectIds.Add(part.ObjectId);
                            }
                        }
                    }
                }
            }

            return arr3dObjectIds;
        }

        public static string GetTagValue(ObjectId idPart)
        {
            string strTagValue = null;

            strTagValue = GetTag(idPart, "Tag");
            if (String.IsNullOrEmpty(strTagValue))
            {
                strTagValue = GetTag(idPart, "LineNumberTag");
            }

            return strTagValue;
        }

        public static PpObjectId GetPpObjectId(ObjectId id)
        {
            PpObjectId ppObjId = new PpObjectId();

            DataLinksManager dlmgr = GetDataLinksManager();
            if (dlmgr != null)
            {
                ppObjId = dlmgr.MakeAcPpObjectId(id);
            }

            return ppObjId;
        }

        private static DataLinksManager GetDataLinksManager()
        {
            DataLinksManager dlmgr = null;

            PlantProject prjRoot = PlantApplication.CurrentProject;
            Project prjPiping = prjRoot.ProjectParts["Piping"];
            if (prjPiping != null && prjPiping.Isloaded())
            {
                dlmgr = prjPiping.DataLinksManager;
            }

            return dlmgr;
        }

        private static string GetTag(ObjectId id, string strTagName)
        {
            string strTag = null;

            DataLinksManager dlmgr = GetDataLinksManager();
            if (dlmgr != null)
            {
                StringCollection arrPropertyNames = new StringCollection();
                arrPropertyNames.Add(strTagName);

                StringCollection arrPropertyValues = dlmgr.GetProperties(id, arrPropertyNames, true);
                if (arrPropertyValues.Count > 0)
                {
                    if (!String.IsNullOrEmpty(arrPropertyValues[0]))
                    {
                        strTag = arrPropertyValues[0];
                    }
                }
            }

            return strTag;
        }
    }
}

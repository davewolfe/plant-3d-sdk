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
using System.Collections;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.ProcessPower.PnIDObjects;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

using Autodesk.ProcessPower.ProjectManager;
using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.DataObjects;


// This is NOT an extension application
[assembly: ExtensionApplication(null)]
// There are commands defined in this class
[assembly: CommandClass(typeof(ListPipelineObjects.Commands))]

namespace ListPipelineObjects
{
    public class Commands
    {
        private static HashSet<ObjectId> m_seachedLineObjIds = new HashSet<ObjectId>();
        private static Project m_prj = null;
        private static DataLinksManager m_dlm = null;

        private static bool Init()
        {
            m_prj = null;
            m_dlm = null;
            try
            {
                m_prj = PlantApplication.CurrentProject.ProjectParts[/*NOXLATE*/"PnId"];

                // 2d Link manager
                //
                m_dlm = m_prj.DataLinksManager;
            }
            catch (System.Exception)
            {
                return false;
            }
            return m_dlm != null;
        }


        [CommandMethod("LISTLINEOBJS", CommandFlags.Modal)]
        public static void ListLineObjects()
        {
            if (!Init())
            {
                return;
            }

            Editor oEditor = AcadApp.DocumentManager.MdiActiveDocument.Editor;

            PromptEntityOptions oPromptOptions = new PromptEntityOptions("Select an object");
            PromptEntityResult  oPromptResult = oEditor.GetEntity(oPromptOptions);
            if (oPromptResult.Status != PromptStatus.OK)
            {
                return;
            }

            m_seachedLineObjIds.Clear();

            Database oDB = AcadApp.DocumentManager.MdiActiveDocument.Database;
            TransactionManager oTxMgr = oDB.TransactionManager;

            StringBuilder sb = new StringBuilder();
            using (Transaction oTx = oTxMgr.StartTransaction())
            {
                DBObject obj = oTx.GetObject(oPromptResult.ObjectId, OpenMode.ForRead);

                LineSegment oLS = obj as LineSegment;
                if (oLS != null)
                {
                    List<Entity> entities = GetEntitiesByLineSegment(oTx, oLS);
                    String lineNumberStr = GetLineNumber(oLS.ObjectId);

                    sb.Append(String.Format("[Line Group: {0}]", lineNumberStr));

                    foreach (Entity lo in entities)
                    {
                        DBObject oDbObj = oTx.GetObject(lo.ObjectId, OpenMode.ForRead);
                        String tagTemplate = GetTagFormat(oDbObj.ObjectId);

                        LineSegment ls = oDbObj as LineSegment;
                        if (ls != null)
                        {
                            String classDispName = GetDisplayName(ls.ClassName);

                            string strMsg = String.Empty;
                            if (String.IsNullOrEmpty(tagTemplate))
                            {
                                strMsg = classDispName;
                            }
                            else
                            {
                                strMsg = String.Format("{0} – {1}", classDispName, ls.TagValue);
                            }

                            sb.Append(Environment.NewLine);
                            sb.Append(strMsg);

                            continue;
                        }

                        Asset oAsset = oDbObj as Asset;
                        if (oAsset != null)
                        {
                            String classDispName = GetDisplayName(oAsset.ClassName);
                            string strMsg = String.Empty;
                            if (String.IsNullOrEmpty(tagTemplate))
                            {
                                strMsg = classDispName;
                            }
                            else
                            {
                                strMsg = String.Format("{0} – {1}", classDispName, oAsset.TagValue);
                            }
                            sb.Append(Environment.NewLine);
                            sb.Append(strMsg);

                            continue;
                        }
#if DEBUG
                        sb.Append(Environment.NewLine);
                        sb.Append("Unknow Object" + oDbObj.GetType().ToString());
#endif
                    }

                    ListPipelineObjectsUIHelpler.ShowWindow(sb.ToString());
                }
                oTx.Commit();
            }
        }

        private static List<Entity> GetEntitiesByLineSegment(Transaction oTx, LineSegment oLS)
        {
            List<Entity> entities = new List<Entity>();

            if (!m_seachedLineObjIds.Contains(oLS.ObjectId))
            {
                m_seachedLineObjIds.Add(oLS.ObjectId);
            }
            else // Just return, if the oLs already been searched.
            {
                return entities;
            }

            // Line objects are metadata classes that represent "something"
            // on the line. They are in order from the start to end of the 
            // line.
            //
            LineObjectCollection los = oLS.GetLineObjectCollection(null);

            bool bIsStart = true;
            foreach (LineObject lo in los)
            {
                if (lo is InlineGap ||
                    lo is InlineIntersection)
                {
                    continue;
                }

                if (lo is EndlineObject)
                {
                    EndlineObject elo = lo as EndlineObject;
                    // Check the begin of the LineObjectCollection to see whether there is other entity
                    List<Entity> beginEntities = CheckAssetForLineObjectConnection(oTx, oLS, elo, bIsStart);
                    AppendEntities(entities, beginEntities);

                    bIsStart = false;
                    continue;
                }
                bIsStart = false;

                DBObject oDbObj = oTx.GetObject(lo.ObjectId, OpenMode.ForRead);

                LineSegment ls = oDbObj as LineSegment;
                if (ls != null)
                {
                    AppendEntity(entities, ls);
                    // If Line  Object on the selected Line Segment is also a LineSegment，
                    // then we also show the line objects on it.
                    List<Entity> entitiesOnLs = GetEntitiesByLineSegment(oTx, ls);
                    AppendEntities(entities, entitiesOnLs);
                    continue;
                }

                Asset oAsset = oDbObj as Asset;
                if (oAsset != null)
                {
                    AppendEntity(entities, oAsset);
                    continue;
                }
            }

            return entities;
        }

        /// <summary>
        /// Check line object at both ends of LineObjectCollection.
        /// </summary>
        /// <param name="bBegin">Specify which end of the LineObjectCollection to check.</param>
        private static List<Entity> CheckAssetForLineObjectConnection(Transaction oTx, LineSegment originalSelectedLine, EndlineObject lo, bool bBegin)
        {
            List<Entity> entities = new List<Entity>();

            DBObject oDbObj = oTx.GetObject(lo.ObjectId, OpenMode.ForRead);
            Asset oAsset = oDbObj as Asset;
            if (oAsset != null)
            {
                List<Entity> segments = CheckLineSegmentForAsset(oTx, originalSelectedLine, oAsset);
                List<Entity> asset = CheckAssetForAsset(oTx, originalSelectedLine, oAsset);

                if (bBegin)
                {
                    AppendEntities(entities, segments);
                    AppendEntity(entities, oAsset); // the asset at end itself
                    AppendEntities(entities, asset);
                }
                else
                {
                    AppendEntities(entities, asset);
                    AppendEntity(entities, oAsset);// the asset at end itself
                    AppendEntities(entities, segments);
                }
            }

            return entities;
        }

        ///
        /// Check other line segment for the Asset.
        ///
        private static List<Entity> CheckLineSegmentForAsset(Transaction oTx, LineSegment originalSelectedLine, Asset oAsset)
        {
            List<Entity> entities = new List<Entity>();

            // Check other line segment for the Asset.
            ObjectIdCollection lineIds = oAsset.LineSegmentIds;
            foreach (ObjectId objId in lineIds)
            {
                if (objId.Equals(originalSelectedLine.ObjectId))
                {
                    continue;
                }

                DBObject assetDbObj = oTx.GetObject(objId, OpenMode.ForRead);

                LineSegment ls = assetDbObj as LineSegment;
                if (ls != null)
                {
                    // If the LineSegment has the same line group number as originalSelectedLine
                    // Then we should also get LineObjects on it.
                    String lsLineNumber = GetLineNumber(ls.ObjectId);
                    String oriLineNumber = GetLineNumber(originalSelectedLine.ObjectId);

                    if (String.Compare(lsLineNumber, oriLineNumber, true) == 0)
                    {
                        List<Entity> beginEntities = GetEntitiesByLineSegment(oTx, ls);
                        AppendEntities(entities, beginEntities);
                    }
                    continue;
                }
            }

            return entities;
        }

        private static List<Entity> CheckAssetForAsset(Transaction oTx, LineSegment originalSelectedLine, Asset oAsset)
        {
            List<Entity> entities = new List<Entity>();

            // Check Owned Assets.
            ObjectIdCollection assetIds = oAsset.OwnedAssets;

            foreach (ObjectId objId in assetIds)
            {
                DBObject assetDbObj = oTx.GetObject(objId, OpenMode.ForRead);

                Asset asset = assetDbObj as Asset;
                if (asset != null)
                {
                    // Check whether the asset of owned by the asset on the originalSelectedLine line segment
                    // has relation ship with originalSelectedLine line segment
                    ObjectIdCollection assetLineIds = oAsset.LineSegmentIds;
                    bool bFound = false;
                    foreach (ObjectId assetLineId in assetLineIds)
                    {
                        if (assetLineId.Equals(originalSelectedLine.ObjectId))
                        {
                            bFound = true;
                            break;
                        }
                    }
                    if (bFound)
                    {
                        AppendEntity(entities, asset);
                    }

                    continue;
                }
            }
            return entities;
        }

        /// <summary>
        /// Append a range to the source list, checking the end of the source Entity against the begin 
        /// of the append list, so that reduntant entity will be ignored.
        /// </summary>
        private static void AppendEntities(List<Entity> sourceList, List<Entity> appendedList)
        {
            if (sourceList.Count > 0 && appendedList.Count > 0)
            {
                Entity lastEntity = sourceList[sourceList.Count - 1];
                Entity appendFirstEntity = appendedList[0];
                if (lastEntity.ObjectId.Equals(appendFirstEntity.ObjectId))
                {
                    sourceList.RemoveAt(sourceList.Count - 1);
                }
            }
            sourceList.AddRange(appendedList);
        }

        /// <summary>
        /// Append an Entity to the source list, checking the end of the source Entity against the 
        /// the appended entity, so that reduntant entity will be ignored.
        /// </summary>
        private static void AppendEntity(List<Entity> sourceList, Entity appendEntity)
        {
            if (sourceList.Count > 0)
            {
                Entity lastEntity = sourceList[sourceList.Count - 1];
                if (lastEntity.ObjectId.Equals(appendEntity.ObjectId))
                {
                    sourceList.RemoveAt(sourceList.Count - 1);
                }
            }
            sourceList.Add(appendEntity);
        }

        private static String GetDisplayName(String className)
        {
            // Find the display name
            String displayName = className;
            
            do{
                PnPDatabase pPnPDb = m_dlm.GetPnPDatabase();

                try{ // Because we start from PnPBase, some tables may not have DisplayNameAttribute
                    displayName = pPnPDb.Tables[className].GetTableAttributeValue(PnPDatabase.DisplayNameAttribute);
                }
                catch (System.Exception)
                {
                    displayName = className;
                }

            }while (false);

            return displayName;
        }

        private static String GetLineNumber(ObjectId objId)
        {
            String lineNumberStr = "Unknow line number";
            do
            {
                int lineGroupId = -1;
                getLineGroupRowId(ref lineGroupId, objId);

                // Get the class name from the entity id
                // string strClassName = pDLM.GetObjectClassname(lineGroupId);
                // Get the table related to the class name
                PnPDatabase db = m_dlm.GetPnPDatabase();
                if (db == null)
                    break;

                PnPTable table = db.Tables["PipeLineGroup"];
                if (table == null)
                    break;

                PnPRow[] rows = table.Select(String.Format("PnPID={0}", lineGroupId));
                if (rows.Length > 0)
                {
                    lineNumberStr = rows[0]["LineNumber"].ToString();
                }

            } while (false);

            return lineNumberStr;
        }

        private static String GetTagFormat(ObjectId objid)
        {
            int rowId = m_dlm.FindAcPpRowId(objid);
            return GetTagFormat(m_dlm, rowId);
        }

        private static String GetTagFormat(DataLinksManager dlm, int rowId)
        {
            if (rowId == -1)
                return String.Empty;

            // Get the class name from the entity id
            string strClassName = m_dlm.GetObjectClassname(rowId);

            // Get the table related to the class name
            PnPDatabase db = m_dlm.GetPnPDatabase();
            if (db == null)
                return string.Empty;

            PnPTable table = db.Tables[strClassName];
            if (table == null)
                return string.Empty;

            return table.GetTableAttributeValue(/*NOXLATE*/"TagFormatName");
        }

        public static void getLineGroupRowId(ref int LineGroupRowId,
                                                ObjectId LineId)
        {
            LineGroupRowId = -1;
            int ridLine = m_dlm.FindAcPpRowId(LineId);
            getLineGroupRowId(ref LineGroupRowId, ridLine);
        }

        public static void getLineGroupRowId(ref int LineGroupRowId,
                                                int LineRowId)
        {
            LineGroupRowId = -1;

            // Go through all the pipelinegroups except the group of the current 
            // line
            String strRelationshipType = /*NOXLATE*/"PipeLineGroupRelationship";
            String strRole2 = /*NOXLATE*/"PipeLineGroup";
            String strRole1 = /*NOXLATE*/"PipeLine";

            StringCollection ClassHeirarchy = new StringCollection();
            String ClassName = m_dlm.GetObjectClassname(LineRowId);
            GetClassHierarchy(out ClassHeirarchy, ClassName);

            if (ClassHeirarchy.Contains(/*NOXLATE*/"SignalLines"))
            {
                strRelationshipType = /*NOXLATE*/"SignalLineGroupRelationship";
                strRole2 = /*NOXLATE*/"SignalLineGroup";
                strRole1 = /*NOXLATE*/"SignalLine";
            }

            PnPRowIdArray arrRowIds = m_dlm.GetRelatedRowIds(strRelationshipType, strRole1, LineRowId, strRole2);
            if (arrRowIds.Count > 0)
            {
                LineGroupRowId = arrRowIds.First.Value; // use only the first value, do not need other values
            }
        }


        public static void GetClassHierarchy(out StringCollection clsHierarchy,
                                                 String className)
        {
            clsHierarchy = new StringCollection();
            try
            {
                PnPTable table = m_dlm.GetPnPDatabase().Tables[className];
                if (table == null) return;

                //bool retval = false;
                while (table != null)
                {
                    if (table.Name == /*NOXLATE*/"PnPBase")
                        break;

                    clsHierarchy.Add(table.Name);
                    table = table.BaseTable;
                }
            }
            catch
            {
                //Debug.Assert(false, e.Message);
            }
        }

    }
}

//
// (C) Copyright 2011-2012 by Autodesk, Inc.
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
using System.Linq;
using System.Text;
using System.Collections.Specialized;
using System.Collections;

// Platform
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

// Plant
using Autodesk.ProcessPower.PnP3dObjects;
using Autodesk.ProcessPower.AcPp3dObjectsUtils;
using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.ProjectManager;
using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.PnP3dDataLinks;
using Autodesk.ProcessPower.P3dProjectParts;
using Autodesk.ProcessPower.PartsRepository;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using PlantApp = Autodesk.ProcessPower.PlantInstance.PlantApplication;
using P3dObjNS = Autodesk.ProcessPower.PnP3dObjects;

[assembly: Autodesk.AutoCAD.Runtime.ExtensionApplication(null)]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(CreatePipeline.Program))]

namespace CreatePipeline
{
    public class Program
    {
        [CommandMethod("CreatePipeline")]
        public static void CreatePipeline()
        { 
            ///////////////////////////////////////////
            /*  MAIN PROCESS
             * 1 Fetch spec part with filter from SPEC (mapping)
             *  a) the filter info comes from pcf
             *  b) assume the spec DB is the one converted from AutoPlant 
             * 2 Create part entity (pipe/fitting/connector and so on) and set entity's props(from spec part)
             * 3 Set part entity's position & orientation
             * 4 Add entity into DB (dwg + dlm)
             * 5 connect entity with the neighbours 
             * 
             * 
             * 
             * complex case: multi branches and loops
             * 
             * big concern: the component order inside pcf file does not strictly follow the pipe routing order
             */
            ///////////////////////////////////////////

            Database db                     = AcadApp.DocumentManager.MdiActiveDocument.Database;
            Editor ed                       = AcadApp.DocumentManager.MdiActiveDocument.Editor;
            ContentManager cm               = ContentManager.GetContentManager();
            Project currentProject          = PlantApp.CurrentProject.ProjectParts["Piping"];
            DataLinksManager dlm            = currentProject.DataLinksManager;
            DataLinksManager3d dlm3d        = DataLinksManager3d.Get3dManager(dlm);
            PipingObjectAdder pipeObjAdder  = new PipingObjectAdder(dlm3d, db);
            Point3d nextPartPos             = Point3d.Origin; // We will start pipe-routing from Origin.

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Create PIPE 1                                                                                          //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //                                                                                                        //
            //                                                                                                        //
            //                                                                                                        //
            //                                                                                                        //
            //     ____________                                                                                       //
            //    |            |                                                                                      //
            //    |            |                                                                                      //
            //    |____________|                                                                                      //
            //                                                                                                        //
            //                                                                                                        //
            //                                                                                                        //
            //                                                                                                        //
            //                                                                                                        //
            //                                                                                                        //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            
            // Set filters to find the correct spec part. For now, we just need to find one 6" (NomDiameter value) pipe.
            //
            StringCollection pipePropertyNames = new StringCollection();
            StringCollection pipePropertyValues = new StringCollection();
            pipePropertyNames.Add("NominalDiameter");
            pipePropertyValues.Add(NomDiameter.ToString());

            // Fetch spec-part PIPE 1
            //
            SpecPart specPart_Pipe1 = FetchSpecPart("Pipe", pipePropertyNames, pipePropertyValues);

            // Create part entity PIPE 1
            //
            Pipe pipeEntity_Pipe1 = CreatePipe();

            // Set part entity position PIPE1
            pipeEntity_Pipe1.StartPoint = nextPartPos;
            pipeEntity_Pipe1.EndPoint = nextPartPos.Add(new Vector3d(60, 0, 0));
            pipeEntity_Pipe1.OuterDiameter = (double)specPart_Pipe1.PropValue("MatchingPipeOd");

            // Add the PIPE 1 part to the database
            //
            ObjectId objId_Pipe1 = AddObjectToDatabase(specPart_Pipe1, pipeEntity_Pipe1, String.Empty, null, ref db, ref ed, ref currentProject, ref dlm, ref dlm3d, ref pipeObjAdder);

            // Fetch the postion of the NEXT part
            //
            nextPartPos = FetchNextPartsPosition(objId_Pipe1, 1, db);

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Create WELD CONNECTOR 1                                                                                //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //                                                                                                        //
            //                                                                                                        //
            //                                                                                                        //
            //                                                                                                        //
            //     ____________                                                                                       //
            //    |            |                                                                                      //
            //    |            |*                                                                                     //
            //    |____________|                                                                                      //
            //                                                                                                        //
            //                                                                                                        //
            //                                                                                                        //
            //                                                                                                        //
            //                                                                                                        //
            //                                                                                                        //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////

            // Fetch spec-part WELD 1
            //
            PartSizePropertiesCollection connectionPropColl_Weld1 = FetchConnectorSpecParts(pipePropertyNames, pipePropertyValues, "Buttweld");

            // Create connector entity WELD 1
            //
            Connector connectorEntity_Weld1 = CreateConnector(connectionPropColl_Weld1, ref db, ref cm, ref currentProject);

            // Set connector entity position WELD 1
            //
            connectorEntity_Weld1.Position = nextPartPos;
            connectorEntity_Weld1.SetOrientation(new Vector3d(1, 0, 0), new Vector3d(0, 0, 1));

            // Add the WELD 1 part to the database
            //
            ObjectId objId_Weld1 = AddObjectToDatabase(null, connectorEntity_Weld1, "Buttweld", connectionPropColl_Weld1, ref db, ref ed, ref currentProject, ref dlm, ref dlm3d, ref pipeObjAdder);

            //  Connect PIPE 1 & WELD 1
            //
            ConnectParts(objId_Pipe1, "S2", objId_Weld1, "S1", db, currentProject);

            // Fetch the postion of the NEXT part
            //
            nextPartPos = FetchNextPartsPosition(objId_Weld1, 1, db);

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Create FLANGE 1                                                                                        //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //                                                                                                        //
            //                                                                                                        //
            //                       __                                                                               //
            //                      /  |                                                                              //
            //     ____________   _/   |                                                                              //
            //    |            | |     |                                                                              //
            //    |            |*|     |                                                                              //
            //    |____________| |_    |                                                                              //
            //                     \   |                                                                              //
            //                      \__|                                                                              //
            //                                                                                                        //
            //                                                                                                        //
            //                                                                                                        //
            //                                                                                                        //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////

            // Fetch spec-part FLANGE 1
            // We will use a different set of properties to fetch the part we want.
            //
            StringCollection flangePropertyNames = new StringCollection();
            StringCollection flangePropertyValues = new StringCollection();
            flangePropertyNames.Add("NominalDiameter");
            flangePropertyValues.Add(NomDiameter.ToString());
            flangePropertyNames.Add("Facing");
            flangePropertyValues.Add("RF");
            flangePropertyNames.Add("PressureClass");
            flangePropertyValues.Add("300");
            SpecPart specPart_Flange1 = FetchSpecPart("Flange", flangePropertyNames, flangePropertyValues);
            ObjectId symbolobjId_Flange1 = cm.GetSymbol(specPart_Flange1, db);

            // Create part Endity FLANGE 1
            //
            PipeInlineAsset partEntity_Flange1 = CreateInlineAsset();

            // Set  the FLANGE 1 part's position & orientation
            //
            partEntity_Flange1.Position = nextPartPos;
            partEntity_Flange1.SetOrientation(new Vector3d(-1, 0, 0), new Vector3d(0, 0, 1));
            partEntity_Flange1.SymbolId = symbolobjId_Flange1;

            // Move the flange
            // We set the X orientation of flange to -1,0,0. This causes the flange     |  We need to move the flange so that pipe ends exactly at Port 2 instead of Port 1.
            // to be flipped (as shown below). The flange is placed at the end          |  To do this, we need to move the flange by a distance equal to Port 2 - Port 1.
            // of the pipe i.e. Port 1 of the flange is placed at the end of the pipe.  |  The function AccomodateParts will calculate and move the part for us.
            //                                                    
            //                           Our (Rotated)                                             Our (Rotated) Flange:             
            //                                     ___                                                              ___             
            //                                    /   |                                                            /   |            
            //            _______________________/____|                                      _________________ ___/    |            
            //                               |        |                                                      ||        |            
            //                         Port2*|        |*Port1                                          Port2*||        |*Port1          
            //            ___________________|________|                                      ________________||___     |            
            //                                   \    |                                                           \    |            
            //                                    \___|                                                            \___|            
            //                                                    
            P3dObjNS.PortCollection partPorts_Flange1 = partEntity_Flange1.GetPorts(PortType.Static);           
            partEntity_Flange1.Position += partPorts_Flange1[0].Position - partPorts_Flange1[1].Position;  

            // Add the FLANGE 1 to the database
            //
            ObjectId objId_Flange1 = AddObjectToDatabase(specPart_Flange1, partEntity_Flange1, String.Empty, null, ref db, ref ed, ref currentProject, ref dlm, ref dlm3d, ref pipeObjAdder);
            
            //  Connect WELD 1 & FLANGE 1 (Note that we are connecting Weld's second port with Flange's second port.)
            //
            ConnectParts(objId_Flange1, "S2", objId_Weld1, "S2", db, currentProject);

            // Fetch the postion of the NEXT part
            //
            nextPartPos = FetchNextPartsPosition(objId_Flange1, 0, db);

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Create FLANGED CONNECTOR 1 (INCLUDES GASKET & BOLTSET)                                                 //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //                                                                                                        //
            //                          |                                                                             //
            //                       __ |                                                                             //
            //                      /  ||                                                                             //
            //     ____________   _/   ||                                                                             //
            //    |            | |     ||                                                                             //
            //    |            |*|     ||                                                                             //
            //    |____________| |_    ||                                                                             //
            //                     \   ||                                                                             //
            //                      \__||                                                                             //
            //                          |                                                                             //
            //                          |                                                                             //
            //                         (_)                                                                            //
            //                                                                                                        //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////

            // Create new filter criteria for flanged connections.
            //
            StringCollection flangedPropertyNames = new StringCollection();
            StringCollection flangedPropertyValues = new StringCollection();
            flangedPropertyNames.Add("NominalDiameter");
            flangedPropertyValues.Add(NomDiameter.ToString());
            flangedPropertyNames.Add("Facing");
            flangedPropertyValues.Add("RF");
            flangedPropertyNames.Add("PressureClass");
            flangedPropertyValues.Add("300");

            // Fetch spec-part GASKET 1
            //
            PartSizePropertiesCollection connectionPropColl_Flanged1 = FetchConnectorSpecParts(flangedPropertyNames, flangedPropertyValues, "Flanged");

            //Create connector entity GASKET 1 
            Connector connectorEntity_Flanged1 = CreateConnector(connectionPropColl_Flanged1, ref db, ref cm, ref currentProject);
            connectorEntity_Flanged1.Position = nextPartPos;
            connectorEntity_Flanged1.SetOrientation(new Vector3d(1, 0, 0), new Vector3d(0, 0, 1));

            // Add the GASKET to the database
            //
            ObjectId objId_Flanged1 = AddObjectToDatabase(null, connectorEntity_Flanged1, "Flanged", connectionPropColl_Flanged1, ref db, ref ed, ref currentProject, ref dlm, ref dlm3d, ref pipeObjAdder);

            //  Connect GASKET 1 & FLANGE 1
            //
            ConnectParts(objId_Flange1, "S1", objId_Flanged1, "S1", db, currentProject);

            // Fetch the postion of the NEXT part
            //
            nextPartPos = FetchNextPartsPosition(objId_Flanged1, 1, db);

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Create FLANGE 2                                                                                        //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //                                                                                                        //
            //                          |                                                                             //
            //                       __ | __                                                                          //
            //                      /  |||  \                                                                         //
            //     ____________   _/   |||   \_                                                                       //
            //    |            | |     |||     |                                                                      //
            //    |            |*|     |||     |                                                                      //
            //    |____________| |_    |||    _|                                                                      //
            //                     \   |||   /                                                                        //
            //                      \__|||__/                                                                         //
            //                          |                                                                             //
            //                          |                                                                             //
            //                         (_)                                                                            //
            //                                                                                                        //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////

            // Fetch spec-part FLANGE 2
            //
            SpecPart specpart_Flange2 = FetchSpecPart("Flange", flangePropertyNames, flangePropertyValues);
            ObjectId symbolobjId_Flange2 = cm.GetSymbol(specpart_Flange2, db);

            // Create part entity FLANGE 2
            //
            PipeInlineAsset partEntity_Flange2 = CreateInlineAsset();

            // Set the FLANGE 2 part's position & orientation
            //
            partEntity_Flange2.Position = nextPartPos;
            partEntity_Flange2.SetOrientation(new Vector3d(1, 0, 0), new Vector3d(0, 0, 1));
            partEntity_Flange2.SymbolId = symbolobjId_Flange2;

            // Add the FLANGE 2 to the database
            //
            ObjectId objId_Flange2 = AddObjectToDatabase(specpart_Flange2, partEntity_Flange2, String.Empty, null, ref db, ref ed, ref currentProject, ref dlm, ref dlm3d, ref pipeObjAdder);

            //  Connect GASKET 1 & FLANGE 2
            //
            ConnectParts(objId_Flange2, "S1", objId_Flanged1, "S2", db, currentProject);

            // Fetch the postion of the NEXT part
            //
            nextPartPos = FetchNextPartsPosition(objId_Flange2, 1, db);

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Create WELD CONNECTOR 2                                                                                //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //                                                                                                        //
            //                          |                                                                             //
            //                       __ | __                                                                          //
            //                      /  |||  \                                                                         //
            //     ____________   _/   |||   \_                                                                       //
            //    |            | |     |||     |                                                                      //
            //    |            |*|     |||     |*                                                                     //
            //    |____________| |_    |||    _|                                                                      //
            //                     \   |||   /                                                                        //
            //                      \__|||__/                                                                         //
            //                          |                                                                             //
            //                          |                                                                             //
            //                         (_)                                                                            //
            //                                                                                                        //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////

            // Fetches spec-part WELD 2
            //
            PartSizePropertiesCollection connectionPropColl_Weld2 = FetchConnectorSpecParts(pipePropertyNames, pipePropertyValues, "Buttweld");

            //Create connector entity WELD 2
            //
            Connector connectorEntity_Weld2 = CreateConnector(connectionPropColl_Weld2, ref db, ref cm, ref currentProject);
            connectorEntity_Weld2.Position = nextPartPos;
            connectorEntity_Weld2.SetOrientation(new Vector3d(1, 0, 0), new Vector3d(0, 0, 1));

            // Add the WELD to the database
            //
            ObjectId objId_Weld2 = AddObjectToDatabase(null, connectorEntity_Weld2, "Buttweld", connectionPropColl_Weld2, ref db, ref ed, ref currentProject, ref dlm, ref dlm3d, ref pipeObjAdder);

            //  Connect PIPE & WELD - CONNECTOR
            //
            ConnectParts(objId_Flange2, "S2", objId_Weld2, "S1", db, currentProject);

            // Fetch the postion of the NEXT part
            //
            nextPartPos = FetchNextPartsPosition(objId_Weld2, 1, db);

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Create PIPE 2                                                                                          //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //                                                                                                        //
            //                          |                                                                             //
            //                       __ | __                                                                          //
            //                      /  |||  \                                                                         //
            //     ____________   _/   |||   \_   ____________                                                        //
            //    |            | |     |||     | |            |                                                       //
            //    |            |*|     |||     |*|            |                                                       //
            //    |____________| |_    |||    _| |____________|                                                       //
            //                     \   |||   /                                                                        //
            //                      \__|||__/                                                                         //
            //                          |                                                                             //
            //                          |                                                                             //
            //                         (_)                                                                            //
            //                                                                                                        //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////

            // Create spec-part PIPE 2 
            //
            SpecPart specpart_Pipe2 = FetchSpecPart("Pipe", pipePropertyNames, pipePropertyValues);

            // Create part entity PIPE 2
            //
            Pipe pipeEntity_Pipe2 = CreatePipe();
            pipeEntity_Pipe2.StartPoint = nextPartPos;
            pipeEntity_Pipe2.EndPoint = nextPartPos.Add(new Vector3d(100, 0, 0));
            pipeEntity_Pipe2.OuterDiameter = (double)specpart_Pipe2.PropValue("MatchingPipeOd");

            // Add the PIPE 2 part to the database
            //
            ObjectId objId_Pipe2 = AddObjectToDatabase(specpart_Pipe2, pipeEntity_Pipe2, String.Empty, null, ref db, ref ed, ref currentProject, ref dlm, ref dlm3d, ref pipeObjAdder);

            //  Connect PIPE 2 & WELD 2
            //
            ConnectParts(objId_Pipe2, "S1", objId_Weld2, "S2", db, currentProject);

            // Fetch the postion of the NEXT part
            //
            nextPartPos = FetchNextPartsPosition(objId_Pipe2, 1, db);

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Create WELD CONNECTOR 3                                                                                //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //                                                                                                        //
            //                          |                                                                             //
            //                       __ | __                                                                          //
            //                      /  |||  \                                                                         //
            //     ____________   _/   |||   \_   ____________                                                        //
            //    |            | |     |||     | |            |                                                       //
            //    |            |*|     |||     |*|            |*                                                      //
            //    |____________| |_    |||    _| |____________|                                                       //
            //                     \   |||   /                                                                        //
            //                      \__|||__/                                                                         //
            //                          |                                                                             //
            //                          |                                                                             //
            //                         (_)                                                                            //
            //                                                                                                        //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////

            // Fetches spec-part WELD 3
            //
            PartSizePropertiesCollection connectionPropColl_Weld3 = FetchConnectorSpecParts(pipePropertyNames, pipePropertyValues, "Buttweld");

            //Create connector entity WELD 3 
            //
            Connector connectorEntity_Weld3 = CreateConnector(connectionPropColl_Weld3, ref db, ref cm, ref currentProject);
            connectorEntity_Weld3.Position = nextPartPos;
            connectorEntity_Weld3.SetOrientation(new Vector3d(1, 0, 0), new Vector3d(0, 0, 1));

            // Add the WELD 3 part to the database
            //
            ObjectId objId_Weld3 = AddObjectToDatabase(null, connectorEntity_Weld3, "Buttweld", connectionPropColl_Weld3, ref db, ref ed, ref currentProject, ref dlm, ref dlm3d, ref pipeObjAdder);

            //  Connect PIPE 2 & WELD 3
            //
            ConnectParts(objId_Pipe2, "S2", objId_Weld3, "S1", db, currentProject);

            // Fetch the postion of the NEXT part
            //
            nextPartPos = FetchNextPartsPosition(objId_Weld3, 1, db);

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Create FLANGE 3                                                                                        //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //                                                                                                        //
            //                          |                                                                             //
            //                       __ | __                        __                                                //
            //                      /  |||  \                      /  |                                               //
            //     ____________   _/   |||   \_   ____________   _/   |                                               //
            //    |            | |     |||     | |            | |     |                                               //
            //    |            |*|     |||     |*|            |*|     |                                               //
            //    |____________| |_    |||    _| |____________| |_    |                                               //
            //                     \   |||   /                    \   |                                               //
            //                      \__|||__/                      \__|                                               //
            //                          |                                                                             //
            //                          |                                                                             //
            //                         (_)                                                                            //
            //                                                                                                        //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////

            // Fetch the FLANGE 3 spec-part
            //
            SpecPart specpart_Flange3 = FetchSpecPart("Flange", flangePropertyNames, flangePropertyValues);
            ObjectId symbolobjId_Flange3 = cm.GetSymbol(specpart_Flange3, db);

            // Create part entity FLANGE 3
            //
            PipeInlineAsset partEntity_Flange3 = CreateInlineAsset();

            // Set FLANGE 3 part's position & orientation
            //
            partEntity_Flange3.Position = nextPartPos;
            partEntity_Flange3.SetOrientation(new Vector3d(-1, 0, 0), new Vector3d(0, 0, 1));
            partEntity_Flange3.SymbolId = symbolobjId_Flange3;
            //moves the flange
            P3dObjNS.PortCollection partPorts_Flange3 = partEntity_Flange3.GetPorts(PortType.Static);
            partEntity_Flange3.Position += partPorts_Flange3[0].Position - partPorts_Flange3[1].Position;

            // Add the FLANGE 3 to the database
            //
            ObjectId objId_Flange3 = AddObjectToDatabase(specpart_Flange3, partEntity_Flange3, String.Empty, null, ref db, ref ed, ref currentProject, ref dlm, ref dlm3d, ref pipeObjAdder);
                        
            //  Connect WELD 3 & FLANGE 3
            //
            ConnectParts(objId_Flange3, "S2", objId_Weld3, "S2", db, currentProject);

            // Fetch the postion of the NEXT part
            //
            nextPartPos = FetchNextPartsPosition(objId_Flange3, 0, db);

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Create FLANGED CONNECTOR 2 (INCLUDES GASKET & BOLTSET)                                                 //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //                                                                                                        //
            //                          |                              |                                              //
            //                       __ | __                        __ |                                              //
            //                      /  |||  \                      /  ||                                              //
            //     ____________   _/   |||   \_   ____________   _/   ||                                              //
            //    |            | |     |||     | |            | |     ||                                              //
            //    |            |*|     |||     |*|            |*|     ||                                              //
            //    |____________| |_    |||    _| |____________| |_    ||                                              //
            //                     \   |||   /                    \   ||                                              //
            //                      \__|||__/                      \__||                                              //
            //                          |                              |                                              //
            //                          |                              |                                              //
            //                         (_)                            (_)                                             //
            //                                                                                                        //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////

            // Create the GASKET 2 part (it fetches the spec-part internally)
            //
            PartSizePropertiesCollection connectionPropColl_Flanged2 = FetchConnectorSpecParts(flangedPropertyNames, flangedPropertyValues, "Flanged");
            Connector connectorEntity_Flanged2 = CreateConnector(connectionPropColl_Flanged2, ref db, ref cm, ref currentProject);
            connectorEntity_Flanged2.Position = nextPartPos;
            connectorEntity_Flanged2.SetOrientation(new Vector3d(1, 0, 0), new Vector3d(0, 0, 1));

            // Add the GASKET 2 to the database
            //
            ObjectId objId_Flanged2 = AddObjectToDatabase(null, connectorEntity_Flanged2, "Flanged", connectionPropColl_Flanged2, ref db, ref ed, ref currentProject, ref dlm, ref dlm3d, ref pipeObjAdder);

            //  Connect FLANGE 3 & GASKET 2
            //
            ConnectParts(objId_Flange3, "S1", objId_Flanged2, "S1", db, currentProject);

            // Fetch the postion of the NEXT part
            //
            nextPartPos = FetchNextPartsPosition(objId_Flanged2, 1, db);

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Create VALVE 1                                                                                         //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //                                                                                                        //
            //                          |                              |    <<<<==>>>>                                //
            //                       __ | __                        __ | ____   ||   ____                             //
            //                      /  |||  \                      /  |||    \  ||  /    |                            //
            //     ____________   _/   |||   \_   ____________   _/   |||     \ || /     |                            //
            //    |            | |     |||     | |            | |     |||      \--/      |                            //
            //    |            |*|     |||     |*|            |*|     |||                |                            //
            //    |____________| |_    |||    _| |____________| |_    |||      /--\      |                            //
            //                     \   |||   /                    \   |||     /    \     |                            //
            //                      \__|||__/                      \__|||____/      \____|                            //
            //                          |                              |                                              //
            //                          |                              |                                              //
            //                         (_)                            (_)                                             //
            //                                                                                                        //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            
            // We need an additional filter to find the correct valve in the spec. So we are going to use the PartFamilyLongDesc
            // property as an additional filter item.
            //
            StringCollection valvePropertyNames = new StringCollection();
            StringCollection valvePropertyValues = new StringCollection();
            valvePropertyNames.Add("NominalDiameter");
            valvePropertyValues.Add(NomDiameter.ToString());
            valvePropertyNames.Add("PartFamilyLongDesc");
            valvePropertyValues.Add("Gate Valve, Solid Wedge, 300 LB, RF, ASME B16.10, ASTM A216 Gr WPB, Hand Wheel");

            // Fetch spec-part VALVE 1
            //
            SpecPart specpart_Valve1 = FetchSpecPart("Valve", valvePropertyNames, valvePropertyValues);
            ObjectId symbolobjId_Valve1 = cm.GetSymbol(specpart_Valve1, db);

            // Create part entity VALVE 1 
            //
            PipeInlineAsset pipeInlineAsset_Valve1 = CreateInlineAsset();

            // Set the VALVE 1 part's position & orientation
            //
            pipeInlineAsset_Valve1.Position = nextPartPos;
            pipeInlineAsset_Valve1.SetOrientation(new Vector3d(1, 0, 0), new Vector3d(0, 0, 1));
            pipeInlineAsset_Valve1.SymbolId = symbolobjId_Valve1;

            // Add the VALVE 1 to the database
            //
            ObjectId objId_Valve1 = AddObjectToDatabase(specpart_Valve1, pipeInlineAsset_Valve1, String.Empty, null, ref db, ref ed, ref currentProject, ref dlm, ref dlm3d, ref pipeObjAdder);

            //  Connect VALVE 1 & GASKET 2
            //
            ConnectParts(objId_Valve1, "S1", objId_Flanged2, "S2", db, currentProject);

            // Fetch the postion of the NEXT part
            //
            nextPartPos = FetchNextPartsPosition(objId_Valve1, 1, db);

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Create FLANGED CONNECTOR 3 (INCLUDES GASKET & BOLTSET)                                                 //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //                                                                                                        //
            //                          |                              |    <<<<==>>>>    |                           //
            //                       __ | __                        __ | ____   ||   ____ |                           //
            //                      /  |||  \                      /  |||    \  ||  /    ||                           //
            //     ____________   _/   |||   \_   ____________   _/   |||     \ || /     ||                           //
            //    |            | |     |||     | |            | |     |||      \--/      ||                           //
            //    |            |*|     |||     |*|            |*|     |||                ||                           //
            //    |____________| |_    |||    _| |____________| |_    |||      /--\      ||                           //
            //                     \   |||   /                    \   |||     /    \     ||                           //
            //                      \__|||__/                      \__|||____/      \____||                           //
            //                          |                              |                  |                           //
            //                          |                              |                  |                           //
            //                         (_)                            (_)                (_)                          //
            //                                                                                                        //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////

            // Create the FLANGED CONNECTOR 3 part
            //
            PartSizePropertiesCollection connectionPropColl_Flanged3 = FetchConnectorSpecParts(flangedPropertyNames, flangedPropertyValues, "Flanged");
            Connector connectorEntity_Flanged3 = CreateConnector(connectionPropColl_Flanged3, ref db, ref cm, ref currentProject);
            connectorEntity_Flanged3.Position = nextPartPos;
            connectorEntity_Flanged3.SetOrientation(new Vector3d(1, 0, 0), new Vector3d(0, 0, 1));

            // Add the GASKET 3 to the database
            //
            ObjectId objId_Flanged3 = AddObjectToDatabase(null, connectorEntity_Flanged3, "Flanged", connectionPropColl_Flanged3, ref db, ref ed, ref currentProject, ref dlm, ref dlm3d, ref pipeObjAdder);

            //  Connect VALVE 1 & GASKET 3
            //
            ConnectParts(objId_Valve1, "S2", objId_Flanged3, "S1", db, currentProject);

            // Fetch the postion of the NEXT part
            //
            nextPartPos = FetchNextPartsPosition(objId_Flanged3, 1, db);

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Create FLANGE 4                                                                                        //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //                                                                                                        //
            //                          |                              |    <<<<==>>>>    |                           //
            //                       __ | __                        __ | ____   ||   ____ | __                        //
            //                      /  |||  \                      /  |||    \  ||  /    |||  \                       //
            //     ____________   _/   |||   \_   ____________   _/   |||     \ || /     |||   \_                     //
            //    |            | |     |||     | |            | |     |||      \--/      |||     |                    //
            //    |            |*|     |||     |*|            |*|     |||                |||     |                    //
            //    |____________| |_    |||    _| |____________| |_    |||      /--\      |||    _|                    //
            //                     \   |||   /                    \   |||     /    \     |||   /                      //
            //                      \__|||__/                      \__|||____/      \____|||__/                       //
            //                          |                              |                  |                           //
            //                          |                              |                  |                           //
            //                         (_)                            (_)                (_)                          //
            //                                                                                                        //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////

            // Fetch spec-part FLANGE 4
            //
            SpecPart specpart_Flange4 = FetchSpecPart("Flange", flangePropertyNames, flangePropertyValues);
            ObjectId symbolobjId_Flange4 = cm.GetSymbol(specpart_Flange4, db);

            // Create part entity FLANGE 4 
            //
            PipeInlineAsset partEntity_Flange4 = CreateInlineAsset();

            // Set  the FLANGE 4 part's orientation
            //
            partEntity_Flange4.Position = nextPartPos;
            partEntity_Flange4.SetOrientation(new Vector3d(1, 0, 0), new Vector3d(0, 0, 1));
            partEntity_Flange4.SymbolId = symbolobjId_Flange4;

            // Add the FLANGE 4 to the database
            //
            ObjectId objId_Flange4 = AddObjectToDatabase(specpart_Flange4, partEntity_Flange4, String.Empty, null, ref db, ref ed, ref currentProject, ref dlm, ref dlm3d, ref pipeObjAdder);

            //  Connect GASKET 3 & FLANGE 4
            //
            ConnectParts(objId_Flange4, "S1", objId_Flanged3, "S2", db, currentProject);

            // Fetch the postion of the NEXT part
            //
            nextPartPos = FetchNextPartsPosition(objId_Flange4, 1, db);

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Create WELD CONNECTOR 4                                                                                //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //                                                                                                        //
            //                          |                              |    <<<<==>>>>    |                           //
            //                       __ | __                        __ | ____   ||   ____ | __                        //
            //                      /  |||  \                      /  |||    \  ||  /    |||  \                       //
            //     ____________   _/   |||   \_   ____________   _/   |||     \ || /     |||   \_                     //
            //    |            | |     |||     | |            | |     |||      \--/      |||     |                    //
            //    |            |*|     |||     |*|            |*|     |||                |||     |*                   //
            //    |____________| |_    |||    _| |____________| |_    |||      /--\      |||    _|                    //
            //                     \   |||   /                    \   |||     /    \     |||   /                      //
            //                      \__|||__/                      \__|||____/      \____|||__/                       //
            //                          |                              |                  |                           //
            //                          |                              |                  |                           //
            //                         (_)                            (_)                (_)                          //
            //                                                                                                        //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////

            // Fetches spec-part WELD 4
            //
            PartSizePropertiesCollection connectionPropColl_Weld4 = FetchConnectorSpecParts(pipePropertyNames, pipePropertyValues, "Buttweld");

            //Create part entity WELD 4 
            //
            Connector connectorEntity_Weld4 = CreateConnector(connectionPropColl_Weld4, ref db, ref cm, ref currentProject);
            connectorEntity_Weld4.Position = nextPartPos;
            connectorEntity_Weld4.SetOrientation(new Vector3d(1, 0, 0), new Vector3d(0, 0, 1));

            // Add the WELD 4 to the database
            //
            ObjectId objId_Weld4 = AddObjectToDatabase(null, connectorEntity_Weld4, "Buttweld", connectionPropColl_Weld4, ref db, ref ed, ref currentProject, ref dlm, ref dlm3d, ref pipeObjAdder);

            //  Connect FLANGE 4 & WELD 4
            //
            ConnectParts(objId_Flange4, "S2", objId_Weld4, "S1", db, currentProject);

            // Fetch the postion of the NEXT part
            //
            nextPartPos = FetchNextPartsPosition(objId_Weld4, 1, db);

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Create PIPE 3                                                                                          //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //                                                                                                        //
            //                          |                              |    <<<<==>>>>    |                           //
            //                       __ | __                        __ | ____   ||   ____ | __                        //
            //                      /  |||  \                      /  |||    \  ||  /    |||  \                       //
            //     ____________   _/   |||   \_   ____________   _/   |||     \ || /     |||   \_   ____________      //
            //    |            | |     |||     | |            | |     |||      \--/      |||     | |            |     //
            //    |            |*|     |||     |*|            |*|     |||                |||     |*|            |     //
            //    |____________| |_    |||    _| |____________| |_    |||      /--\      |||    _| |____________|     //
            //                     \   |||   /                    \   |||     /    \     |||   /                      //
            //                      \__|||__/                      \__|||____/      \____|||__/                       //
            //                          |                              |                  |                           //
            //                          |                              |                  |                           //
            //                         (_)                            (_)                (_)                          //
            //                                                                                                        //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////


            // Create the PIPE 3 spec-part
            //
            SpecPart specPart_Pipe3 = FetchSpecPart("Pipe", pipePropertyNames, pipePropertyValues);

            // Create the PIPE 3 part
            //
            Pipe pipeEntity_Pipe3 = CreatePipe();
            pipeEntity_Pipe3.StartPoint = nextPartPos;
            pipeEntity_Pipe3.EndPoint = nextPartPos.Add(new Vector3d(50, 0, 0));
            pipeEntity_Pipe3.OuterDiameter = (double)specPart_Pipe3.PropValue("MatchingPipeOd");

            // Add the PIPE 3 to the database
            //
            ObjectId objId_Pipe3 = AddObjectToDatabase(specPart_Pipe3, pipeEntity_Pipe3, String.Empty, null, ref db, ref ed, ref currentProject, ref dlm, ref dlm3d, ref pipeObjAdder);

            //  Connect PIPE 3 & WELD 4
            //
            ConnectParts(objId_Pipe3, "S1", objId_Weld4, "S2", db, currentProject); 
        }

        public static Pipe CreatePipe()
        {
            Pipe pipePart = new Pipe();
            return pipePart;
        }

        public static PartSizePropertiesCollection FetchConnectorSpecParts(
            StringCollection propertyNames,
            StringCollection propertyValues,
            String           jointName)
        {
            // For this sample, we are going to use only the "Flanged" and "Buttweld" joints.
            // More joint types can be added here.
            //
            PartSizePropertiesCollection pspColl = new Autodesk.ProcessPower.PnP3dObjects.PartSizePropertiesCollection();
            switch (jointName.ToLower())
            {
                case "flanged":
                    {
                        var spGPart = FetchSpecPart("Gasket", propertyNames, propertyValues);
                        pspColl.Add(spGPart);

                        var spBPart = FetchSpecPart("BoltSet", propertyNames, propertyValues);
                        pspColl.Add(spBPart);

                        break;
                    }
                case "buttweld":
                    {
                        PartSizeProperties psprops = new NonSpecPart();
                        psprops.Name = "Buttweld";
                        psprops.Type = "Buttweld";
                        psprops.NominalDiameter = NomDiameter;
                        psprops.Spec = SpecName;
                        psprops.SetPropValue("NominalDiameter", NomDiameter.Value);
                        psprops.SetPropValue("NominalUnit", NomDiameter.Units);
                        psprops.SetPropValue("PartSizeLongDesc", "Buttweld");
                        psprops.SetPropValue("ShortDescription", "Buttweld");
                        pspColl.Add(psprops);

                        break;
                    }
            }
            return pspColl;
        }

        public static Connector CreateConnector(
            PartSizePropertiesCollection pspColl,
            ref Database db,
            ref ContentManager cm,
            ref Project currentProject)
        {
            Connector connectorPart = new Connector();
            connectorPart.SlopeTolerance = 0.1;
            connectorPart.OffsetTolerance = 0.0;

            foreach (PartSizeProperties psp in pspColl)
            {
                switch (psp.Type.ToLower())
                {
                    case "gasket":
                        {
                            var blockSubPart = new BlockSubPart();
                            try
                            {
                                if (psp.PropValue("ContentGeometryParamDefinition") != null)
                                    blockSubPart.SymbolId = cm.GetSymbol(psp, db);
                                blockSubPart.PartSizeProperties.NominalDiameter = NomDiameter;
                                connectorPart.AddSubPart(blockSubPart);
                            }
                            catch { }
                            break;
                        }
                    case "boltset":
                        {
                            BoltSetSubPart boltsetSubPart = new BoltSetSubPart();
                            boltsetSubPart.PartSizeProperties.NominalDiameter = NomDiameter;
                            boltsetSubPart.Length = NomDiameter.Value * 2;
                            connectorPart.AddSubPart(boltsetSubPart);
                            break;
                        }
                    case "buttweld":
                        {
                            Double dWeldGap = 0;
                            String bWeldGap = String.Empty;
                            currentProject.GetProjectVariable("USEWELDGAPS", out bWeldGap);
                            if (String.Compare("TRUE", bWeldGap, true) == 0)
                            {
                                String sWeldGap = String.Empty;
                                currentProject.GetProjectVariable("WELDGAPSIZE", out sWeldGap);
                                if (!String.IsNullOrEmpty(sWeldGap))
                                {
                                    Double.TryParse(sWeldGap, out dWeldGap);
                                }
                            }

                            WeldSubPart weldSubPart = new WeldSubPart();
                            weldSubPart.Width = dWeldGap;
                            weldSubPart.PartSizeProperties.NominalDiameter = NomDiameter;
                            weldSubPart.PartSizeProperties.Type = "Buttweld";
                            weldSubPart.PartSizeProperties.Spec = SpecName;
                            connectorPart.AddSubPart(weldSubPart);
                            break;
                        }
                }
            }
            return connectorPart;
        }

        public static PipeInlineAsset CreateInlineAsset()
        {
            PipeInlineAsset pipeInlineAssetPart = new PipeInlineAsset();
            return pipeInlineAssetPart;
        }

        public static SpecPart FetchSpecPart(
            String PartName, 
            StringCollection PropertyNames, 
            StringCollection PropertyValues)
        {
            // spec mananger object of the active project
            var specMgr = SpecManager.GetSpecManager();

            SpecPart specPart = null;
            if (specMgr.HasType(SpecName, PartName))
            {
                SpecPartReader specPartReader = specMgr.SelectParts(SpecName, PartName, PropertyNames, PropertyValues);
                while (specPartReader.Next())
                {
                    specPart = specPartReader.Current;
                    var partND = specPart.NominalDiameter.Value;

                    if (specPart.Type.Equals(PartName) &&
                        partND == NomDiameter.Value) // matching part found
                    {
                        break;
                    }
                }
            }

            // save the spectpart object
            //
            return specPart;
        }

        public static ObjectId AddObjectToDatabase(
            SpecPart specPart,
            Autodesk.ProcessPower.PnP3dObjects.Part part,
            String connectionName,
            PartSizePropertiesCollection connectionPropColl,
            ref Database db, 
            ref Editor ed,
            ref Project currentProject,
            ref DataLinksManager dlm,
            ref DataLinksManager3d dlm3d,
            ref PipingObjectAdder pipeObjAdder)
        {
            ObjectId partObjectId = ObjectId.Null;
            if (pipeObjAdder == null)
            {
                ed.WriteMessage("Error: Cannot create PipingObjectAdder");
                return partObjectId;
            }

            Autodesk.AutoCAD.DatabaseServices.TransactionManager tm = db.TransactionManager;
            try
            {
                using (Transaction trans = tm.StartTransaction())
                {
                    if (part.GetType() == typeof(Pipe))
                    {
                        Pipe pipe = part as Pipe;
                        pipeObjAdder.Add(specPart, pipe);
                    }
                    else if (part.GetType() == typeof(PipeInlineAsset))
                    {
                        PipeInlineAsset pipeInlineAsset = part as PipeInlineAsset;
                        pipeObjAdder.Add(specPart, pipeInlineAsset);
                    }
                    else if (part.GetType() == typeof(Connector))
                    {
                        Connector connector = part as Connector;
                        pipeObjAdder.Add(connectionName, connectionPropColl, connector);
                    }

                    partObjectId = part.ObjectId;
                    trans.AddNewlyCreatedDBObject(part, true);
                    trans.Commit();
                }
            }
            catch (SystemException ex)
            {
                ed.WriteMessage("Error: Exception while appending objects.\n");
                ed.WriteMessage(ex.Message);
            }
            return partObjectId;
        }

        public static void ConnectParts(
            ObjectId objIdPipeOrInlineAssetPart, 
            String portnamePipeOrInlineAssetPart,
            ObjectId objIdConnectorPart,
            String portnameConnectorPart,
            Database db, 
            Project currentProject)
        {
            Autodesk.AutoCAD.DatabaseServices.TransactionManager tm = db.TransactionManager;            
            try
            {
                using (Transaction trans = tm.StartTransaction())
                {
                    ConnectionManager cm        = new ConnectionManager();

                    Autodesk.ProcessPower.PnP3dObjects.Part pipeorinlineassetPart  = null;
                    Autodesk.ProcessPower.PnP3dObjects.Part connectorPart         = null;

                    if (objIdPipeOrInlineAssetPart != ObjectId.Null)
                        pipeorinlineassetPart = tm.GetObject(objIdPipeOrInlineAssetPart, OpenMode.ForRead) as Autodesk.ProcessPower.PnP3dObjects.Part;
                    if (objIdConnectorPart != ObjectId.Null)
                        connectorPart = tm.GetObject(objIdConnectorPart, OpenMode.ForRead) as Autodesk.ProcessPower.PnP3dObjects.Part;
                     
                    // Fetch Pipe/InlineAsset ports information
                    //
                    Autodesk.ProcessPower.PnP3dObjects.PortCollection partPorts = pipeorinlineassetPart.GetPorts(PortType.Static);

                    // Fetch Connector ports information
                    //
                    Autodesk.ProcessPower.PnP3dObjects.PortCollection connectorPorts = connectorPart.GetPorts(PortType.Static);

                    // Create port pairs and connect them
                    //
                    Pair partPair = new Pair();
                    partPair.Port = partPorts[portnamePipeOrInlineAssetPart];
                    partPair.ObjectId = pipeorinlineassetPart.ObjectId;

                    Pair conctrPair = new Pair();
                    conctrPair.Port = connectorPorts[portnameConnectorPart];
                    conctrPair.ObjectId = connectorPart.ObjectId;

                    // Connect the parts and save
                    //
                    cm.Connect(partPair, conctrPair);
                    trans.Commit();
                }
            }
            catch (SystemException ex)
            {
                Editor ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage("error: while connecting objects.\n");
                ed.WriteMessage(ex.Message);
            }
        }

        public static void AccomodatePart(
            ObjectId objIdPart,
            bool bFlipped,
            Database db)
        {
            Autodesk.AutoCAD.DatabaseServices.TransactionManager tm = db.TransactionManager;
            try
            {
                using (Transaction trans = tm.StartTransaction())
                {
                    Autodesk.ProcessPower.PnP3dObjects.Part part                = null;
                    Autodesk.ProcessPower.PnP3dObjects.PortCollection partPorts = null;

                    // Fetch Pipe/InlineAsset/Connector ports information
                    //
                    if (objIdPart != ObjectId.Null)
                    {
                        part = tm.GetObject(objIdPart, OpenMode.ForWrite) as Autodesk.ProcessPower.PnP3dObjects.Part;
                        partPorts = part.GetPorts(PortType.Static);
                    }

                    // Calculate how muct the part should be moved to match ports
                    //
                    if (partPorts != null && partPorts.Count >= 2)
                    {
                        Vector3d moveVector;
                        if(bFlipped)
                            moveVector = partPorts[0].Position - partPorts[1].Position;
                        else
                            moveVector = partPorts[1].Position - partPorts[0].Position;

                        if (part != null)
                        {
                            Point3d oldPos = part.Position;
                            part.Position = oldPos.Add(moveVector);
                        }
                    }

                    trans.Commit();
                }
            }
            catch (SystemException ex)
            {
                Editor ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage("error: while connecting objects.\n");
                ed.WriteMessage(ex.Message);
            }
        }

        public static Point3d FetchNextPartsPosition(
            ObjectId objIdPart,
            int iPortNum,
            Database db)
        {
            Autodesk.AutoCAD.DatabaseServices.TransactionManager tm = db.TransactionManager;
            try
            {
                using (Transaction trans = tm.StartTransaction())
                {
                    Autodesk.ProcessPower.PnP3dObjects.Part part = null;
                    Autodesk.ProcessPower.PnP3dObjects.PortCollection partPorts = null;

                    // Fetch Pipe/InlineAsset/Connector ports information
                    //
                    if (objIdPart != ObjectId.Null)
                    {
                        part = tm.GetObject(objIdPart, OpenMode.ForRead) as Autodesk.ProcessPower.PnP3dObjects.Part;
                        partPorts = part.GetPorts(PortType.Static);
                        if (partPorts.Count > iPortNum)
                            return partPorts[iPortNum].Position;
                    }

                }
            }
            catch (SystemException ex)
            {
                Editor ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage("error: while connecting objects.\n");
                ed.WriteMessage(ex.Message);
            }

            return Point3d.Origin;
        }

        static String SpecName
        {
            get
            {
                return "CS300";
            }
        }

        static NominalDiameter NomDiameter
        {
            get
            {
                return new NominalDiameter("in", 6.0);
            }
        }
    }
}

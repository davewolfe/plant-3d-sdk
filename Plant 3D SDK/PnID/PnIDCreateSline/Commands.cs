//
// (C) Copyright 2013 by Autodesk, Inc.
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
using System.Reflection;

using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Windows.Forms;
using Autodesk.AutoCAD.EditorInput;


[assembly: CommandClass(typeof(PnIDCreateSline.Commands))]

namespace PnIDCreateSline
{
    public class Commands
    {
        #region Commands

        [CommandMethod(/*NOXLATE*/"CreateSline_Automatic", CommandFlags.Modal)]
        public static void CreateSline_Automatic()
        {
            CreatePnID.CreateSline_Automatic();
        }

        [CommandMethod(/*NOXLATE*/"CreateSline_Manual", CommandFlags.Modal)]
        public static void CreateSline_Manual()
        {
            CreatePnID.CreateSline_Manual();
        }

        #endregion
    }        
}
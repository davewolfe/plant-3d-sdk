//
// (C) Copyright 2009-2013 by Autodesk, Inc.
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

using Autodesk.AutoCAD.Runtime;
using Autodesk.ProcessPower.Validation;
using Autodesk.ProcessPower.PnIDDwgValidation;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

// There are commands defined in this class
[assembly: CommandClass(typeof(Autodesk.ProcessPower.PipingValidation.PipingValidationCommands))]
// This is NOT an extension application
[assembly: ExtensionApplication(null)]

namespace Autodesk.ProcessPower.PipingValidation
{
    public partial class PipingValidationCommands
    {
        [CommandMethod("PipingValidationCommands", "ADDPIPINGVALIDATIONRULES", "ADDPIPINGVALIDATIONRULES", CommandFlags.Session)]
        public static void AddPipingValidationRules()
        {
            AcPpValidationManager vmgr = ValidationSingleton.Manager;
            if (!vmgr.Rules.ContainsGuid(UnconnectedPortRule.RuleGuid))
            {
                vmgr.Rules.Add(new UnconnectedPortRule());
                vmgr.Settings.EnabledRuleGuids.Add(UnconnectedPortRule.RuleGuid);

                AcPpValidationRuleGroup group = new AcPpValidationRuleGroup();
                group = vmgr.Settings.RuleGroups["Piping.Group.Guid"];
                group.RuleGuids.Add(UnconnectedPortRule.RuleGuid);

                AcadApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage(UnconnectedPortRule.RuleName + " is added to validation mechanism.");
            }
            else
            {
                AcadApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage(UnconnectedPortRule.RuleName + " has already existed in validation mechanism.");
            }
        }
    }
}

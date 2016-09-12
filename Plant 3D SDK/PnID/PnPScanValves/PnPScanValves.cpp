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

// APITest.cpp : Defines the entry point for the DLL application.
//

#include "stdafx.h"
#include "math.h"
#include <map>

#ifdef _MANAGED
#pragma managed(push, off)
#endif

// Function prototypes
//
AcPpDataLinksManager* getDataLinksManager();
void SortValves(AcDbObjectIdArray& scannedValves);
int qsortValves(const void* first, const void* second); 
void AssignTagToValves(const AcDbObjectIdArray& valves);
Acad::ErrorStatus AssignProperty(const AcDbObjectId& target, 
								 const ACHAR* pszValveCode, 
								 const ACHAR* pszValveNumber); 
AcString GetValveCode(const AcString& className); 
void PrintReport(const AcDbObjectIdArray& scannedValves);
Acad::ErrorStatus LoadLineSegments(AcDbObjectIdArray& arrLineSegment); 
bool isAssetValve(const AcDbObjectId& isValve);


/// <summary>
/// This function gets the data links manager using the project database's 
/// full path name that is obtained from the PnId services.
/// </summary>
AcPpDataLinksManager* getDataLinksManager()
{
	AcPpDataLinksManager* pLm = NULL;	// return value

    TCHAR projectName[4096];			// MAX_PATH is often defined as 4096
    projectName[0] = _T('\0');

    AcPnIDServices* pServices = AcPnIDServices::getPnIDServices();
    if (pServices != NULL)
    {
        pServices->getDataLinksManagerName(projectName, 4096);
		EPpDLStatus dlStat = AcPpDataLinksManager::GetManager(projectName, pLm);
		if (dlStat != ePpDLOk)
		{
			acutPrintf(L"\nError obtaining data links manager for the current PnId project.");
		}
    }

	return pLm;
}


/// <summary>
/// This function scans through all the valves (under the class "HandValves") sorts, and tags them. 
/// To tag them, it asks for the starting number by the user. 
/// </summary>
void ScanValves()
{
	Acad::ErrorStatus es = Acad::eOk; 
	AcDbObjectIdArray scannedValves; 
	AcDbObjectIdArray arrLineSegment;

	acutPrintf(L"\n-----------------------------------------------------");
	acutPrintf(L"\n Scan for valves in the drawing");
	acutPrintf(L"\n-----------------------------------------------------");

	// Get a list of SLINE's in current space; this is done to show
	// how the objects on an SLINE are iterated to find items of interest.
	//
	es = LoadLineSegments(arrLineSegment); 
	if (es == Acad::eOk)
	{
		for(int i = 0; i < arrLineSegment.length(); i++)
		{
			acutPrintf(L"\nLineSegment %d is %x: ", i+1, arrLineSegment.at(i)); 
			
			AcDbObjectPointer<AcPpLineSegment> spLine(arrLineSegment[i], AcDb::kForRead);		 
			if (spLine.openStatus() == Acad::eOk)
			{
				AcPpLineObjectIterator *it = spLine->lineObjectIterator(NULL);
				for (it->start(); !it->done(); it->step())
				{                                        
					AcPpLineObject *pLineObj = it->object();	
					AcDbObjectId oValveId = pLineObj->objectId();	

					// Item of interest?
					//
					if (isAssetValve(oValveId))
					{
						scannedValves.append(oValveId);						
					}		
				}
				delete it;
			}
			else
			{
				acutPrintf(L"\nCould not open SLINE for read!");
			}
		}

		if(!scannedValves.isEmpty())
		{
			SortValves(scannedValves); 
			acutPrintf(L"\n Sorting complete");

			AssignTagToValves(scannedValves);
			acutPrintf(L"\n Tags assigned to Valves"); 

			acutPrintf(L"\n-----------------------------------------------------");
			acutPrintf(L"\nScanning ended");
			acutPrintf(L"\n-----------------------------------------------------");
		}
		else
		{
			acutPrintf(L"\n No Valves found"); 
		}
	}
	else
	{
		acutPrintf(L"\n Line segments not loaded. Command failed"); 
	}
}

/// <summary>
///  This function returns true if and only if the asset is of class "HandValves"
/// </summary>
bool isAssetValve(const AcDbObjectId& isValve)
{
	wchar_t strClassName[256];
    strClassName[0] = _T('\0');
	bool itIs = false; 

    AcPpDataLinksManager* pLm = getDataLinksManager();
	if (pLm != NULL)
	{
		pLm->GetObjectClassname(isValve, strClassName, 256); 
		if(pLm->IsKindOf(strClassName, _T("HandValves"))) 		
		{
			itIs = true; 							
			acutPrintf(L"\n Valve found. It is: %s", strClassName);
		}
		else 
		{
			itIs = false; 
		}
	}
	else
	{
		acutPrintf(L"\nError determining valve type.");
	}

	return itIs; 
}

/// <summary>
/// This function sorts all the valves present in the drawing. 
/// The top right corner of the screen is taken as reference. First the Y co-ordinates of the valves are considered. 
/// Positions closer to the top right corner are given preference. If Y co-ordinates are the same, X co-ordinates are considered. 
/// Closer the X co-ordinate to the top right, greater is the valve preference. 
/// </summary>
void SortValves(AcDbObjectIdArray& scannedValves)
{
	acutPrintf(L"\nIn SortValves function \n\n"); 	
	acutPrintf(L"\n Before sorting: \n"); 
	PrintReport(scannedValves);
	
	qsort(&scannedValves[0], scannedValves.length(), sizeof(scannedValves[0]), qsortValves);
	
	acutPrintf(L"\n After sorting: \n"); 
	PrintReport(scannedValves);
}

/// <summary>
///  Prints report of the valve positions. 
/// </summary>
void PrintReport(const AcDbObjectIdArray& scannedValves)
{
	for (int l = 0; l < scannedValves.length(); l++)
	{
		AcDbObjectId valveId = scannedValves.at(l); 

		AcDbObjectPointer<AcPpAsset> spAsset(valveId, AcDb::kForRead);
		if (spAsset.openStatus() == Acad::eOk)
		{
			AcGePoint3d oPosition = spAsset->position();
			acutPrintf(L"\nPosition of asset %s is: (%f, %f, %f)", 
						spAsset->className().constPtr(), 
					    oPosition.x, oPosition.y, oPosition.z);
		}	
		else
		{
			acutPrintf(L"\nError in opening entity %d.", l); 
		}
	}
}

/// <summary>
/// This function assigns tags to assets belonging to "HandValves" class. 
/// For example, a GateValve is tagged as GV-104. Here the part before the 
/// hyphen depends on the "class" of the Valve and the part after the hyphen 
/// depends on the valve position. 
/// </summary>
void AssignTagToValves(const AcDbObjectIdArray& valves)
{
	int iStartIndex = 0; 
	int iRetCode = acedGetInt(L"\nEnter starting valve number: ", &iStartIndex);
	if (iRetCode != RTNORM)
	{
		acutPrintf(L"\n Command canceled. No tags assigned"); 
		return;
	}

	if (iStartIndex < 0)
	{
		acutPrintf(L"\n Incorrect value entered. Default start value of 100 assigned");
		iStartIndex = 100;
	}

	for (int i = 0; i < valves.length(); i++)
	{
		AcDbObjectPointer<AcPpAsset> spAsset(valves.at(i), AcDb::kForRead);
		if (spAsset.openStatus() == Acad::eOk)
		{
			// Valve code
			//
			AcString className = spAsset->className(); 
			AcString valveCode = GetValveCode(className);
			if (valveCode.isEmpty())
			{
				valveCode = "Valve"; 
			}

			// Valve number
			//
			AcString valveNumber;
			valveNumber.format(L"%d", iStartIndex+i);

			Acad::ErrorStatus es = AssignProperty(valves.at(i), valveCode, valveNumber);
			if (es == Acad::eOk)
			{
				acutPrintf(L"\n The valve %s has been tagged as %s-%s. ", 
					className.constPtr(), valveCode.constPtr(), 
					valveNumber.constPtr());
			}
		}
		else
		{
			acutPrintf(L"\nError in opening entity %d", i);
		}
	}	
}

/// <summary>
/// Assigns Tag to the valves. Called from AssignTagToValves() function. 
/// </summary>
Acad::ErrorStatus AssignProperty(const AcDbObjectId& target,
								 const ACHAR* pszValveCode,
								 const ACHAR* pszValveNumber)
{
	if (pszValveCode == NULL || pszValveNumber == NULL)
	{
		return Acad::eNullPtr;
	}
	if (target.isNull())
	{
		return Acad::eNullObjectId;
	}

	AcPpDataLinksManager* pLm = getDataLinksManager();
	if (pLm == NULL)
	{
		return Acad::eNotHandled;
	}

	AcPpStringArray propNames;
	AcPpStringArray propVals;

	propNames.Append(L"Code");
	propNames.Append(L"Number");
	
	propVals.Append(pszValveCode);
	propVals.Append(pszValveNumber);

	EPpDLStatus dlStat = pLm->SetProperties(target, propNames, propVals);

	return dlStat == ePpDLOk ? Acad::eOk : Acad::eNotHandled;
}

/// <summary>
/// Sorts valves depending on their position. For algorithm see the SortValves documentation. 
/// Uses standard qsort algorithm. 
/// </summary>
int qsortValves(const void* first, const void* second)
{
	AcDbObjectId valveId1 = *(AcDbObjectId*)first; 
	AcDbObjectId valveId2 = *(AcDbObjectId*)second;
	double tol = 1e-6; 
	int retValue = 0;

	AcDbObjectPointer<AcPpAsset> spAsset1(valveId1, AcDb::kForRead, true);
	AcDbObjectPointer<AcPpAsset> spAsset2(valveId2, AcDb::kForRead, true);
	if (spAsset1.openStatus() != Acad::eOk ||
		spAsset2.openStatus() != Acad::eOk)
	{
		return 0;
	}

	AcGeVector3d diff = spAsset1->position().asVector() - spAsset2->position().asVector(); 

	double withY = diff.dotProduct(AcGeVector3d::kYAxis); //positive value: spAsset2 is belows pAsset1
	double withX = diff.dotProduct(AcGeVector3d::kXAxis); //positive value: spAsset2 is to the left of spAsset1
	
	if (withY < -1*tol) 
	{
		retValue = 1; 
	}
	else if (withY > tol)
	{
		retValue = -1; 
	}
	else if (fabs(withY - tol) < 1e-4)
	{
		if(withX < -1*tol)
		{
			retValue = 1; 
		}
		else if (withX > tol)
		{
			retValue = -1; 
		}
		else
		{
			retValue = 0; 
		}
	}

	return retValue; 
}

/// <summary>
/// Loads all the line segments present in the drawing. 
/// The reason we load line segments is because we assume that all the valves are directly connected to a line segment. Always. 
/// </summary>
Acad::ErrorStatus LoadLineSegments(AcDbObjectIdArray& arrLineSegment)
{
	const AcDbDatabase* pDb = acdbHostApplicationServices()->workingDatabase();
	AcDbObjectId  btrId;
	if (pDb != NULL)
	{
		btrId = pDb->currentSpaceId();
	}

	Acad::ErrorStatus es = Acad::eOk;
    if(btrId.isNull())
	{
        return Acad::eInvalidInput;
	}

	arrLineSegment.removeAll();

    AcDbBlockTableRecordPointer pBTR(btrId, AcDb::kForRead);
	if ((es = pBTR.openStatus()) == Acad::eOk)
	{
		AcDbBlockTableRecordIterator *pIter = NULL;
		if ((es = pBTR->newIterator(pIter)) == Acad::eOk)
		{
			for (pIter->start(); !pIter->done(); pIter->step()) 
			{
				AcDbObjectId entId;
				es = pIter->getEntityId(entId);

				if (es == Acad::eOk && 
					entId.objectClass()->isDerivedFrom(AcPpLineSegment::desc()))
				{
					arrLineSegment.append(entId);
				}
			}
			delete pIter;
		}
	}
    return es;
}

/// <summary>
/// Map of Valve classname and the code. Thess codes have been given by the developer. No specific standards were followed. 
/// </summary>
AcString GetValveCode(const AcString& className)
{
	static bool isMapInit = false;
	static std::map<AcString, AcString> myStrMap; 
	
	if (!isMapInit)
	{
		isMapInit = true;

		myStrMap[L"PressureAndVacuumReliefValve"] = L"PVRV"; 
		myStrMap[L"ThreeWayBallValve"] = L"TWBV"; 
		myStrMap[L"ThreeWayGlobeValve"] = L"TWGV";
		myStrMap[L"AngleValve"] = L"AV"; 
		myStrMap[L"AngleBallValve"] = L"ABV";
		myStrMap[L"AngleGlobeValve"] = L"AGV"; 
		myStrMap[L"AngleSpringSafetyValve"] = L"ASSV";
		myStrMap[L"SafetyValve"] = L"SV"; 
		myStrMap[L"AngleSafetyValve"] = L"ASV";
		myStrMap[L"PressureReliefValve"] = L"PRV"; 
		myStrMap[L"VacuumReliefValve"] = L"VRV";
		myStrMap[L"RuptureDiskForPressureRelief"] = L"RDPR";
		myStrMap[L"NeedleValve"] = L"NV"; 
	
		myStrMap[L"PinchValve"] = L"PnV";
		myStrMap[L"RotaryValve"] = L"RV";
		myStrMap[L"FourWayValve"] = L"FWV";
		myStrMap[L"DiverterValve"] = L"DV";
		myStrMap[L"CheckValve"] = L"ChV";
		myStrMap[L"GenericRotaryValve"] = L"GRV"; 

		myStrMap[L"StopCheckValve"] = L"SCV";
		myStrMap[L"DiaphragmValve"] = L"DpV";
		myStrMap[L"ExcessFlowValve"] = L"EFV";
		myStrMap[L"KnifeValve"] = L"KV";
		myStrMap[L"ThreeWayValve"] = L"TWV";
		myStrMap[L"ButterflyValve"] = L"BfV"; 

		myStrMap[L"ControlValve"] = L"CV";
		myStrMap[L"BallValve"] = L"BV";
		myStrMap[L"PlugValve"] = L"PV";
		myStrMap[L"GlobeValve"] = L"GbV";
		myStrMap[L"GateValve"] = L"GtV";
		myStrMap[L"BallValveClosed"] = L"BVC"; 

		myStrMap[L"GateValveClosed"] = L"GtVC";
		myStrMap[L"GlobeValveClosed"] = L"GbVC";
		myStrMap[L"NeedleValveClosed"] = L"NVC";
		myStrMap[L"PlugValveClosed"] = L"PVC";
	}
	
	AcString str; 
	std::map<AcString, AcString>::iterator iter = myStrMap.find(className.constPtr());
	if (iter != myStrMap.end())
	{
		str = iter->second; 
	}
	
	return str;
}

extern "C"
AcRx::AppRetCode 
acrxEntryPoint(AcRx::AppMsgCode msg, void* pkt)
{
	switch(msg)
	{
		case AcRx::kInitAppMsg:
			acrxRegisterAppMDIAware(pkt);
			acrxDynamicLinker->unlockApplication(pkt);
			
			acedRegCmds->addCommand(
				L"PNIDAPITEST_GROUP", L"SCANVALVES", L"SCANVALVES", 
				ACRX_CMD_MODAL | ACRX_CMD_USEPICKSET, 
				ScanValves); 

            // Register RX classes
			acrxBuildClassHierarchy();
			break;
		case AcRx::kUnloadAppMsg:
            // Unregister RX classes
			acedRegCmds->removeGroup(L"PNIDAPITEST_GROUP");
			break;
	}

	return AcRx::kRetOK;
}


BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
					 )
{
    return TRUE;
}


#ifdef _MANAGED
#pragma managed(pop)
#endif


(C) Copyright 2014 by Autodesk, Inc.

Equipment Plant SDK sample Readme

1. Build
- Open Equipment.csproj in DevStudio
- Add Reference Paths for AutoCAD and Plant dlls
- Build to produce Equipment.dll

2. Launch
- Launch AutoCAD Plant 3D application
- Run command NETLOAD and select Equipment.dll

3. Commands
- EquipmentLoadPackage loads equipment type from the content and sets current 
- EquipmentLoadTemplate loads equipment type from the template in the current project and sets current
- EquipmentLoadEntity loads equipment from the entity and sets current both equipment type and entity
- EquipmentCreate creates new entity with the current type
- EquipmentConvert converts ACAD entities into new equipment, and sets current both type and entity
- EquipmentModify edits current equipment type, and if the current entity is set, updates it as well
	Different options for different equipments:
	- Attach/Detach graphics, both need current entity set
	- Shapes for fabricated equipment
	- Parameters for parametric one
- EquipmentSaveTemplate saves curent equipment type as a templete in the current project
- EquipmentAddNozzle adds new default nozzle to fabricated or converted equipment
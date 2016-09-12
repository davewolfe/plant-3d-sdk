(C) Copyright 2014 by Autodesk, Inc.

PipeSupport Plant SDK sample Readme

1. Build
- Open PipeSupport.csproj in DevStudio
- Add Reference Paths for AutoCAD and Plant dlls
- Build to produce PipeSupport.dll

2. Launch
- Launch AutoCAD Plant 3D application
- Run command NETLOAD and select PipeSupport.dll

3. Commands
- SupportScriptCreate places parametric pipe support from existing script     
- SupportSpecCreate places pipe support from PipeSupportSpec
- SupportBlockCreate places custom pipe support from ACAD block
- SupportSelectionCreate places custom pipe support from ACAD entities
- SupportConvert convert ACAD entities into pipe support
- SupportAttach attaches additional entities to support block
- SupportDetach detaches previously attached entities
- SupportEdit edits support parameters and updates its geometry
- SupportRegister registers support script to enable it in Plant.
    Python script must exist in the Plant content (metadata.zip, variants.zip)
    For example, existing unused dummy leg scripts: csgd001_mj, csgd001_of, csgd001_om...
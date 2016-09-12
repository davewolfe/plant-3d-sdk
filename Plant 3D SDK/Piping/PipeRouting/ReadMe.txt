(C) Copyright 2014 by Autodesk, Inc.

PipeRouting Plant SDK sample Readme

1. Build
- Open PipeRouting.csproj in DevStudio
- Add Reference Paths for AutoCAD and Plant dlls
- Build to produce PipeRouting.dll

2. Launch
- Launch AutoCAD Plant 3D application
- Run command NETLOAD and select PipeRouting.dll

3. Commands
- PipeRoute - places pipes, elbows, cutback elbows, bent pipes, FLPs, asymmetric parts, reducers when sizes don't match,
              does branching and auto routing.
- FittingAdd - places fittings and olets
- FittingEdit - edits geometry of parametric fittings. Can be used to edit supports and parametric equipments as well
- FittingSubstitue - substitutes fittings

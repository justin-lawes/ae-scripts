// Rename Comp to Project Name
// Renames the active/selected composition to match the project file name (without extension).
// Run via File > Scripts > Run Script File

(function renameCompToProjectName() {

    try {

        // --- Guard: project must be saved ---
        var projectFile = app.project.file;
        if (!projectFile) {
            alert("Project has not been saved yet.\nSave the project first, then run this script.");
            return;
        }

        // --- Derive name: strip .aep or .aepx extension ---
        var fileName = projectFile.name;
        var newName = fileName.replace(/\.(aepx?)$/i, "");

        // --- Guard: find selected composition ---
        // Prefer app.project.activeItem (frontmost comp viewer); fall back to Project panel selection.
        var targetComp = null;

        // Check active item — valid if a comp viewer is frontmost
        var activeItem = app.project.activeItem;
        if (activeItem instanceof CompItem) {
            targetComp = activeItem;
        }

        // If no active comp, look at Project panel selection
        if (!targetComp) {
            var sel = app.project.selection;
            var selLength = (sel && typeof sel.length === "number") ? sel.length : 0;

            if (selLength === 0) {
                alert("No composition selected.\nSelect a composition in the Project panel or open one in the Composition viewer.");
                return;
            }

            // Filter to CompItems only
            var comps = [];
            for (var i = 0; i < selLength; i++) {
                if (sel[i] instanceof CompItem) {
                    comps.push(sel[i]);
                }
            }

            if (comps.length === 0) {
                alert("No composition in the current selection.\nSelect a composition in the Project panel or open one in the Composition viewer.");
                return;
            }

            if (comps.length > 1) {
                alert("Multiple compositions selected (" + comps.length + ").\nSelect exactly one composition and run the script again.");
                return;
            }

            targetComp = comps[0];
        }

        // Final guard: ensure targetComp is actually a CompItem
        if (!(targetComp instanceof CompItem)) {
            alert("Could not identify a valid composition to rename.");
            return;
        }

        // --- Already named correctly? ---
        if (targetComp.name === newName) {
            alert("Comp is already named \"" + newName + "\". Nothing to do.");
            return;
        }

        // --- Rename inside an undo group ---
        var oldName = targetComp.name;
        app.beginUndoGroup("Rename Comp to Project Name");
        targetComp.name = newName;
        app.endUndoGroup();

    } catch (e) {
        alert("Script error: " + e.toString());
    }

}());

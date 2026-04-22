// RenderFolderPanel.jsx
// Dockable ScriptUI panel for After Effects
// Auto-routes new render queue items to a selected output folder.
//
// INSTALL: Copy to  <AE install>/Scripts/ScriptUI Panels/
//          Then access via  Window > RenderFolderPanel

(function (thisObj) {

    var SETTINGS_SECTION = "RenderFolderPanel";
    var STATE_KEY        = "renderFolderPanelState";

    // ── Persistent global state (survives panel close/reopen in same AE session) ──
    if (!$.global[STATE_KEY]) {
        $.global[STATE_KEY] = {
            selectedFolder  : null,   // fsName string of chosen folder
            isEnabled       : false,
            lastQueueCount  : 0,
            taskID          : null,
            pausedForRender : []      // items temporarily unqueued for targeted render
        };
    }
    var state = $.global[STATE_KEY];

    // ── Shared reroute helpers (in $.global so scheduleTask's eval can reach them) ─

    // Strip characters that are illegal in Windows/Mac filenames.
    $.global.sanitizeFilename = function (name) {
        return name.replace(/[\/\\:*?"<>|]/g, "_");
    };

    // Returns true if filePath is already assigned to an output module whose
    // CURRENT path is not currentFsName (i.e. a different module owns it).
    // We identify "same module" by its existing path because ExtendScript
    // returns a new wrapper object on every outputModules[j] access, making
    // === identity checks unreliable.
    $.global.isOutputPathInUse = function (filePath, currentFsName) {
        try {
            if (!app.project) return false;
            var norm        = filePath.toLowerCase().replace(/\\/g, "/");
            var currentNorm = currentFsName ? currentFsName.toLowerCase().replace(/\\/g, "/") : null;
            var rq          = app.project.renderQueue;
            for (var i = 1; i <= rq.numItems; i++) {
                try {
                    var item = rq.item(i);
                    // DONE items have already rendered — they no longer "own" their path
                    if (item.status === RQItemStatus.DONE) continue;
                    for (var j = 1; j <= item.outputModules.length; j++) {
                        try {
                            var f = item.outputModules[j].file;
                            if (!f) continue;
                            var fNorm = f.fsName.toLowerCase().replace(/\\/g, "/");
                            if (fNorm === currentNorm) continue; // this is the module we're editing
                            if (fNorm === norm) return true;
                        } catch (e) {}
                    }
                } catch (e) {}
            }
        } catch (e) {}
        return false;
    };

    // Always names the file after the comp (no AE-appended "_1" suffixes).
    // If folder is null, keeps the existing directory but still fixes the name.
    // Skips silently if the target path is already owned by another queue item.
    $.global.rerouteOutputModule = function (om, compName, folder) {
        try {
            var currentFile = om.file;
            if (!currentFile) return;
            // Preserve the extension AE chose for this output module
            var extMatch = currentFile.name.match(/(\.[^.]+)$/);
            var ext      = extMatch ? extMatch[1] : "";
            var dir      = folder ? folder : currentFile.parent.fsName;
            var newPath  = dir + "/" + $.global.sanitizeFilename(compName) + ext;
            // Skip if already correct
            if (currentFile.fsName.toLowerCase().replace(/\\/g, "/") ===
                newPath.toLowerCase().replace(/\\/g, "/")) return;
            // Skip if another queue item already owns this path (prevents AE dialog)
            if ($.global.isOutputPathInUse(newPath, currentFile.fsName)) return;
            om.file = new File(newPath);
        } catch (e) {}
    };

    // Delete any existing output files for all QUEUED items so AE never prompts
    // to overwrite — called right before rq.render().
    $.global.deleteExistingOutputFiles = function () {
        try {
            if (!app.project) return;
            var rq = app.project.renderQueue;
            for (var i = 1; i <= rq.numItems; i++) {
                try {
                    var item = rq.item(i);
                    if (item.status !== RQItemStatus.QUEUED) continue;
                    for (var j = 1; j <= item.outputModules.length; j++) {
                        try {
                            var f = item.outputModules[j].file;
                            if (f && f.exists) f.remove();
                        } catch (e) {}
                    }
                } catch (e) {}
            }
        } catch (e) {}
    };

    // Local alias for convenience inside the IIFE
    var rerouteOutputModule = $.global.rerouteOutputModule;

    // ── Scheduled-task callback (must live in global scope) ──────────────────────
    $.global.renderFolderPanelCheck = function () {
        var s = $.global[STATE_KEY];
        if (!s) return;
        try {
            if (!app.project) return;
            var rq = app.project.renderQueue;

            // Restore items that were temporarily unqueued for a targeted render
            if (s.pausedForRender && s.pausedForRender.length > 0 && !rq.rendering) {
                for (var q = 0; q < s.pausedForRender.length; q++) {
                    try { s.pausedForRender[q].render = true; } catch (e) {}
                }
                s.pausedForRender = [];
            }

            if (!s.isEnabled || !s.selectedFolder) return;
            if (rq.rendering) return;  // never touch the queue while AE is rendering
            var currentCount = rq.numItems;

            if (currentCount > s.lastQueueCount) {
                for (var i = s.lastQueueCount + 1; i <= currentCount; i++) {
                    try {
                        var item      = rq.item(i);
                        var compName  = item.comp.name;
                        for (var j = 1; j <= item.outputModules.length; j++) {
                            rerouteOutputModule(item.outputModules[j], compName, s.selectedFolder);
                        }
                    } catch (itemErr) {}
                }
            }
            s.lastQueueCount = currentCount;
        } catch (e) { /* project may not be open */ }
    };

    // ── Settings (persisted across AE restarts) ──────────────────────────────────
    function loadSettings() {
        try {
            if (app.settings.haveSetting(SETTINGS_SECTION, "folder")) {
                var fp = app.settings.getSetting(SETTINGS_SECTION, "folder");
                if (fp && new Folder(fp).exists) state.selectedFolder = fp;
            }
            if (app.settings.haveSetting(SETTINGS_SECTION, "enabled")) {
                state.isEnabled = (app.settings.getSetting(SETTINGS_SECTION, "enabled") === "true");
            }
        } catch (e) {}
    }

    function saveSettings() {
        try {
            if (state.selectedFolder)
                app.settings.saveSetting(SETTINGS_SECTION, "folder", state.selectedFolder);
            app.settings.saveSetting(SETTINGS_SECTION, "enabled", state.isEnabled ? "true" : "false");
        } catch (e) {}
    }

    // ── Queue monitoring ─────────────────────────────────────────────────────────
    function startMonitoring() {
        try { state.lastQueueCount = app.project ? app.project.renderQueue.numItems : 0; } catch (e) {}
        if (state.taskID !== null) {
            try { app.cancelTask(state.taskID); } catch (e) {}
            state.taskID = null;
        }
        state.taskID = app.scheduleTask("renderFolderPanelCheck()", 2000, true);
    }

    function stopMonitoring() {
        if (state.taskID !== null) {
            try { app.cancelTask(state.taskID); } catch (e) {}
            state.taskID = null;
        }
    }

    // ── Add selected comp(s) to queue, apply folder, and render ──────────────────
    function addSelectedAndRender() {
        if (!app.project) { alert("No project is open."); return; }

        // Collect all selected CompItems from the project panel
        var comps = [];
        for (var i = 1; i <= app.project.numItems; i++) {
            var item = app.project.item(i);
            if (item.selected && (item instanceof CompItem)) comps.push(item);
        }

        if (comps.length === 0) {
            alert("No composition selected.\nSelect one or more comps in the Project panel first.");
            return;
        }

        var rq     = app.project.renderQueue;
        var failed = [];

        // Temporarily unqueue everything already in the queue so only our
        // new items render when we trigger the Render button.
        var paused = [];
        for (var q = 1; q <= rq.numItems; q++) {
            try {
                var qi = rq.item(q);
                if (qi.status === RQItemStatus.QUEUED) {
                    qi.render = false;
                    paused.push(qi);
                }
            } catch (e) {}
        }

        for (var c = 0; c < comps.length; c++) {
            try {
                // Use .id comparison — === on AE wrapper objects is unreliable
                for (var k = rq.numItems; k >= 1; k--) {
                    try {
                        var existing = rq.item(k);
                        if (existing.comp.id === comps[c].id &&
                            existing.status !== RQItemStatus.RENDERING) {
                            existing.remove();
                        }
                    } catch (e) {}
                }

                var rqItem   = rq.items.add(comps[c]);
                var compName = comps[c].name;

                for (var j = 1; j <= rqItem.outputModules.length; j++) {
                    rerouteOutputModule(rqItem.outputModules[j], compName, state.selectedFolder);
                }
            } catch (addErr) {
                failed.push(comps[c].name);
            }
        }

        state.lastQueueCount = rq.numItems;

        if (failed.length > 0) {
            for (var q = 0; q < paused.length; q++) {
                try { paused[q].render = true; } catch (e) {}
            }
            alert("Could not add to render queue:\n" + failed.join("\n"));
            return;
        }

        // Store paused items — renderFolderPanelCheck restores them after render.
        state.pausedForRender = paused;

        // Delete any existing output files so AE won't prompt to overwrite.
        $.global.deleteExistingOutputFiles();

        // Open the Render Queue panel.
        app.executeCommand(2161);

        // Trigger AE's native Render button via a background PowerShell script
        // using Windows UI Automation. This goes through AE's proper render path
        // so the progress bar shows — unlike calling rq.render() from script.
        var ps = new File(Folder.temp.fsName + "/ae_render_trigger.ps1");
        ps.open("w");
        ps.writeln("Start-Sleep -Milliseconds 700");
        ps.writeln("Add-Type -AssemblyName UIAutomationClient");
        ps.writeln("Add-Type -AssemblyName UIAutomationTypes");
        ps.writeln("try {");
        ps.writeln("  $root    = [System.Windows.Automation.AutomationElement]::RootElement");
        ps.writeln("  $cond    = New-Object System.Windows.Automation.AndCondition(");
        ps.writeln("    (New-Object System.Windows.Automation.PropertyCondition(");
        ps.writeln("      [System.Windows.Automation.AutomationElement]::ControlTypeProperty,");
        ps.writeln("      [System.Windows.Automation.ControlType]::Button)),");
        ps.writeln("    (New-Object System.Windows.Automation.PropertyCondition(");
        ps.writeln("      [System.Windows.Automation.AutomationElement]::NameProperty, 'Render')))");
        ps.writeln("  $btn = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $cond)");
        ps.writeln("  if ($btn) {");
        ps.writeln("    $btn.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()");
        ps.writeln("  }");
        ps.writeln("} catch {}");
        ps.close();

        var psPath = ps.fsName.replace(/\//g, "\\");
        system.callSystem('cmd /c start "" /b powershell -NoProfile -ExecutionPolicy Bypass -File "' + psPath + '"');
    }

    // ── Apply folder to every eligible item already in the queue ─────────────────
    function applyFolderToAllItems() {
        if (!state.selectedFolder) {
            alert("Please select an output folder first.");
            return;
        }
        if (!new Folder(state.selectedFolder).exists) {
            alert("The selected folder no longer exists.\nPlease choose a new folder.");
            return;
        }
        if (!app.project) { alert("No project is open."); return; }

        var rq    = app.project.renderQueue;
        var count = 0;
        for (var i = 1; i <= rq.numItems; i++) {
            try {
                var item = rq.item(i);
                // Skip items that are actively rendering or already finished
                if (item.status === RQItemStatus.RENDERING ||
                    item.status === RQItemStatus.DONE) continue;

                var compName = item.comp.name;
                for (var j = 1; j <= item.outputModules.length; j++) {
                    rerouteOutputModule(item.outputModules[j], compName, state.selectedFolder);
                }
                count++;
            } catch (e) {}
        }
        alert("Updated " + count + " render queue item" + (count === 1 ? "" : "s") + ".");
    }

    // ── UI ───────────────────────────────────────────────────────────────────────
    function buildUI(thisObj) {
        var panel = (thisObj instanceof Panel)
            ? thisObj
            : new Window("palette", "Render Folder", undefined, { resizeable: true });

        panel.orientation  = "column";
        panel.alignChildren = ["fill", "top"];
        panel.spacing      = 8;
        panel.margins      = [10, 12, 10, 12];

        // ── Heading ──
        var heading = panel.add("statictext", undefined, "RENDER OUTPUT FOLDER");
        heading.alignment = ["center", "top"];

        panel.add("panel", undefined, "").alignment = ["fill", "top"];  // divider

        // ── Folder path row ──
        var folderGroup = panel.add("group");
        folderGroup.orientation   = "column";
        folderGroup.alignChildren = ["fill", "top"];
        folderGroup.alignment     = ["fill", "top"];
        folderGroup.spacing       = 4;

        folderGroup.add("statictext", undefined, "Output Folder:");

        var pathRow = folderGroup.add("group");
        pathRow.alignment     = ["fill", "top"];
        pathRow.alignChildren = ["fill", "center"];
        pathRow.spacing       = 4;

        var pathDisplay = pathRow.add("edittext", undefined, "");
        pathDisplay.alignment     = ["fill", "center"];
        pathDisplay.preferredSize = [-1, 22];
        pathDisplay.helpTip       = "Type or paste a folder path, then press Enter.\nOr use the … button to browse.";

        var browseBtn = pathRow.add("button", undefined, "\u2026");  // …
        browseBtn.preferredSize = [30, 22];
        browseBtn.helpTip = "Choose output folder";

        // ── Enable toggle ──
        var enableRow = panel.add("group");
        enableRow.alignment = ["fill", "top"];
        var enableCheck = enableRow.add("checkbox", undefined, "Auto-route new queue items");
        enableCheck.helpTip = "Monitors the render queue and redirects every\nnewly added item to the folder above.";

        // ── Status indicator ──
        var statusLabel = panel.add("statictext", undefined, "");
        statusLabel.alignment     = ["fill", "top"];
        statusLabel.preferredSize = [-1, 18];

        panel.add("panel", undefined, "").alignment = ["fill", "top"];  // divider

        // ── Add & Render button ──
        var renderBtn = panel.add("button", undefined, "\u25B6  Add Selected Comp & Render");
        renderBtn.alignment     = ["fill", "top"];
        renderBtn.preferredSize = [-1, 28];
        renderBtn.helpTip       = "Adds every selected comp to the render queue,\nroutes output to the folder above, then renders.";

        // ── Apply-to-all button ──
        var applyBtn = panel.add("button", undefined, "Apply Folder to All Queue Items");
        applyBtn.alignment     = ["fill", "top"];
        applyBtn.preferredSize = [-1, 26];
        applyBtn.helpTip       = "Reroute every item currently in the render queue\n(skips items that are rendering or already done).";

        // ── Update helpers ────────────────────────────────────────────────────────
        function updateUI() {
            // Only overwrite the text if the user isn't actively editing it
            if (!pathDisplay.active)
                pathDisplay.text = state.selectedFolder || "No folder selected";
            enableCheck.value = state.isEnabled;
            applyBtn.enabled  = !!state.selectedFolder;

            if (state.isEnabled && state.selectedFolder) {
                statusLabel.text = "\u25CF Monitoring render queue\u2026";
            } else if (state.isEnabled && !state.selectedFolder) {
                statusLabel.text = "\u26A0  Enabled \u2014 no folder set";
            } else {
                statusLabel.text = "\u25CB Monitoring off";
            }
        }

        // ── Event handlers ────────────────────────────────────────────────────────

        // Accept a typed/pasted path when the field loses focus or Enter is pressed.
        function commitTypedPath() {
            var typed = pathDisplay.text;
            if (!typed || typed === "No folder selected") return;
            var f = new Folder(typed);
            if (f.exists) {
                state.selectedFolder = f.fsName;
                updateUI();
                saveSettings();
                if (state.isEnabled) startMonitoring();
            } else {
                // Flash the old value back so the field doesn't stay wrong
                pathDisplay.text = state.selectedFolder || "No folder selected";
                alert("Folder not found:\n" + typed);
            }
        }
        pathDisplay.onDeactivate = commitTypedPath;
        pathDisplay.addEventListener("keydown", function (kd) {
            if (kd.keyName === "Enter") { commitTypedPath(); }
        });

        browseBtn.onClick = function () {
            var start  = state.selectedFolder ? new Folder(state.selectedFolder) : Folder.desktop;
            var chosen = Folder.selectDialog("Select Render Output Folder", start);
            if (chosen) {
                state.selectedFolder = chosen.fsName;
                updateUI();
                saveSettings();
                if (state.isEnabled) startMonitoring();  // reset queue baseline
            }
        };

        enableCheck.onClick = function () {
            state.isEnabled = enableCheck.value;
            updateUI();
            saveSettings();
            if (state.isEnabled) startMonitoring();
            else                 stopMonitoring();
        };

        renderBtn.onClick = function () { addSelectedAndRender(); };
        applyBtn.onClick  = function () { applyFolderToAllItems(); };

        // ── Initial render ────────────────────────────────────────────────────────
        updateUI();

        if (panel instanceof Window) {
            panel.center();
            panel.show();
        } else {
            panel.layout.layout(true);
        }

        return panel;
    }

    // ── Boot ─────────────────────────────────────────────────────────────────────
    loadSettings();
    buildUI(thisObj);

    if (state.isEnabled && state.selectedFolder) startMonitoring();

})(this);

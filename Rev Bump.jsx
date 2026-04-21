// revBump.jsx
// Increments the letter or version number in filenames like:
//   Particles_B_v01.aep  OR  HatchDawn_Particles_A_Navy_v01.aep
// Works as a dockable ScriptUI panel or floating dialog.
//
// Naming convention: Prefix_LETTER[_MiddleName]_vNN
//   e.g.  HatchDawn_Particles_A_Navy_v02
//          └─prefix──┘         └┘ └──┘ └─┘
//                               A  Navy  v02

(function (thisObj) {

    // Matches: prefix_LETTER[_middle]_vNN  (middle segment like "Navy" is optional)
    var PATTERN = /^(.+)_([A-Za-z])_(?:(.+)_)?(v)(\d+)$/;

    // ── Parsing ──────────────────────────────────────────────────────────────

    // proj is app.project captured in the calling event handler.
    // Returns { name, fromFile } or null.
    function getSourceName(proj) {
        // Try project file name first
        var fileName = null;
        try { fileName = proj.file.name; } catch (e) {}
        if (fileName) {
            return { name: fileName.replace(/\.aep$/i, ""), fromFile: true };
        }
        // Fall back: scan all items for a comp matching the pattern
        var numItems = 0;
        try { numItems = proj.numItems; } catch (e) {}
        for (var i = 1; i <= numItems; i++) {
            try {
                var item = proj.item(i);
                if (item instanceof CompItem && PATTERN.test(item.name)) {
                    return { name: item.name, fromFile: false };
                }
            } catch (e) {}
        }
        return null;
    }

    function parseName(base) {
        // Two-step: strip _vNN suffix, then find the letter segment.
        // Avoids an ExtendScript regex bug where group 2 is undefined when
        // the optional group 3 (?:(.+)_)? captures a value in one-pass matching.
        var vMatch = base.match(/_(v\d+)$/i);
        if (!vMatch) return null;
        var vFull   = vMatch[1];
        var vPrefix = vFull.charAt(0);
        var numStr  = vFull.slice(1);
        var body = base.slice(0, base.length - vMatch[0].length);

        // Find prefix_LETTER_middle (non-greedy prefix to get the FIRST _X_ segment)
        var m = body.match(/^(.+?)_([A-Za-z])_(.+)$/);
        if (m) return { prefix: m[1], letter: m[2], middle: m[3], vPrefix: vPrefix, numStr: numStr };

        // Find prefix_LETTER (no middle)
        m = body.match(/^(.+?)_([A-Za-z])$/);
        if (m) return { prefix: m[1], letter: m[2], middle: null, vPrefix: vPrefix, numStr: numStr };

        return null;
    }

    function buildName(parts) {
        var mid = parts.middle ? "_" + parts.middle : "";
        return parts.prefix + "_" + parts.letter + mid + "_" + parts.vPrefix + parts.numStr;
    }

    function incrementLetter(parts) {
        var code = parts.letter.toUpperCase().charCodeAt(0);
        var wrapped = (code >= 90); // Z wraps to A
        var newCode = wrapped ? 65 : code + 1;
        var newLetter = String.fromCharCode(newCode);
        if (parts.letter === parts.letter.toLowerCase()) newLetter = newLetter.toLowerCase();
        var newParts = { prefix: parts.prefix, letter: newLetter, middle: parts.middle, vPrefix: parts.vPrefix, numStr: parts.numStr };
        return { name: buildName(newParts), wrapped: wrapped };
    }

    function incrementNumber(parts) {
        var num = parseInt(parts.numStr, 10) + 1;
        var padLen = Math.max(parts.numStr.length, String(num).length);
        var padded = String(num);
        while (padded.length < padLen) padded = "0" + padded;
        var newParts = { prefix: parts.prefix, letter: parts.letter, middle: parts.middle, vPrefix: parts.vPrefix, numStr: padded };
        return buildName(newParts);
    }

    // ── Save / Rename ─────────────────────────────────────────────────────────

    function renameMatchingComp(proj, oldName, newName) {
        for (var i = 1; i <= proj.numItems; i++) {
            var item = proj.item(i);
            if (item instanceof CompItem && item.name === oldName) {
                item.name = newName;
                break;
            }
        }
    }

    function saveAs(proj, oldBaseName, newBaseName) {
        var dir = proj.file.parent;
        var newFile = new File(dir.fsName + "/" + newBaseName + ".aep");
        if (newFile.exists) {
            if (!confirm('"' + newBaseName + '.aep" already exists. Overwrite?')) return false;
        }
        proj.save(newFile);
        renameMatchingComp(proj, oldBaseName, newBaseName);
        return true;
    }

    // ── UI ───────────────────────────────────────────────────────────────────

    function buildUI(thisObj) {
        var win = (thisObj instanceof Panel)
            ? thisObj
            : new Window("palette", "Rev Bump", undefined, { resizeable: false });

        win.orientation = "column";
        win.alignChildren = ["fill", "top"];
        win.spacing = 8;
        win.margins = 12;

        // Current name
        var grpCurrent = win.add("group");
        grpCurrent.orientation = "row";
        grpCurrent.alignChildren = ["left", "center"];
        grpCurrent.add("statictext", undefined, "Current:");
        var txtCurrent = grpCurrent.add("statictext", undefined, "\u2014");
        txtCurrent.justify = "left";
        txtCurrent.preferredSize.width = 240;

        // Source indicator (grey note shown when reading from comp name)
        var txtSource = win.add("statictext", undefined, "");
        txtSource.justify = "left";
        try {
            txtSource.graphics.foregroundColor = txtSource.graphics.newPen(
                txtSource.graphics.PenType.SOLID_COLOR, [0.55, 0.55, 0.55, 1], 1
            );
        } catch (e) {}

        var sep1 = win.add("panel", undefined, "");
        sep1.alignment = ["fill", "top"];
        sep1.preferredSize.height = 2;

        // Previews
        var grpLetterPreview = win.add("group");
        grpLetterPreview.orientation = "row";
        var lblLetter = grpLetterPreview.add("statictext", undefined, "Letter \u2191:");
        lblLetter.preferredSize.width = 72;
        var txtLetterPreview = grpLetterPreview.add("statictext", undefined, "\u2014");
        txtLetterPreview.justify = "left";
        txtLetterPreview.preferredSize.width = 210;

        var grpNumberPreview = win.add("group");
        grpNumberPreview.orientation = "row";
        var lblNumber = grpNumberPreview.add("statictext", undefined, "Number \u2191:");
        lblNumber.preferredSize.width = 72;
        var txtNumberPreview = grpNumberPreview.add("statictext", undefined, "\u2014");
        txtNumberPreview.justify = "left";
        txtNumberPreview.preferredSize.width = 210;

        // Warning label (Z→A wrap notice)
        var txtWarning = win.add("statictext", undefined, "");
        txtWarning.justify = "center";
        try {
            txtWarning.graphics.foregroundColor = txtWarning.graphics.newPen(
                txtWarning.graphics.PenType.SOLID_COLOR, [0.9, 0.6, 0.1, 1], 1
            );
        } catch (e) {}

        var sep2 = win.add("panel", undefined, "");
        sep2.alignment = ["fill", "top"];
        sep2.preferredSize.height = 2;

        // Buttons
        var grpBtns = win.add("group");
        grpBtns.orientation = "row";
        grpBtns.alignment = ["fill", "top"];
        grpBtns.alignChildren = ["fill", "center"];
        var btnLetter = grpBtns.add("button", undefined, "Letter \u2191");
        var btnNumber = grpBtns.add("button", undefined, "Number \u2191");

        var btnRefresh = win.add("button", undefined, "Refresh");

        // ── State ─────────────────────────────────────────────────────────

        var state = null;

        // proj must be app.project captured fresh in the calling event handler.
        // NOTE: In AE 2026, app.project accessed inside nested functions returns
        // stale/empty data. Passing it explicitly from the onClick handler works
        // around this ExtendScript scoping issue.
        function updateUI(proj) {
            try { _updateUI(proj); } catch (e) { alert("Version Up error: " + e.toString() + " (line " + e.line + ")"); }
        }

        function _updateUI(proj) {
            txtWarning.text = "";

            var source = getSourceName(proj);

            if (!source) {
                var noProj = true;
                try { if (proj && proj.numItems >= 0) noProj = false; } catch (e) {}
                txtCurrent.text = noProj ? "(no project open)" : "(no matching comp found)";
                txtSource.text = "";
                txtLetterPreview.text = "\u2014";
                txtNumberPreview.text = "\u2014";
                btnLetter.enabled = false;
                btnNumber.enabled = false;
                state = null;
                win.layout.layout(true);
                return;
            }

            txtSource.text = source.fromFile ? "" : "reading from active comp";

            var parts = parseName(source.name);

            if (!parts) {
                txtCurrent.text = source.name;
                txtLetterPreview.text = "(no match)";
                txtNumberPreview.text = "(no match)";
                btnLetter.enabled = false;
                btnNumber.enabled = false;
                state = null;
                win.layout.layout(true);
                return;
            }

            var letterResult = incrementLetter(parts);
            var numberName   = incrementNumber(parts);

            txtCurrent.text       = source.name;
            txtLetterPreview.text = letterResult.name;
            txtNumberPreview.text = numberName;
            btnLetter.enabled = true;
            btnNumber.enabled = true;

            if (letterResult.wrapped) txtWarning.text = "Z wrapped to A";

            state = {
                currentName: source.name,
                letterName:  letterResult.name,
                numberName:  numberName,
                fromFile:    source.fromFile,
                proj:        proj
            };

            win.layout.layout(true);
        }

        // ── Button handlers ───────────────────────────────────────────────

        btnLetter.onClick = function () {
            if (!state) return;
            try {
                var ok = state.fromFile
                    ? saveAs(state.proj, state.currentName, state.letterName)
                    : (renameMatchingComp(state.proj, state.currentName, state.letterName), true);
                if (ok) updateUI(app.project);
            } catch (e) {
                alert("Operation failed:\n" + e.toString());
            }
        };

        btnNumber.onClick = function () {
            if (!state) return;
            try {
                var ok = state.fromFile
                    ? saveAs(state.proj, state.currentName, state.numberName)
                    : (renameMatchingComp(state.proj, state.currentName, state.numberName), true);
                if (ok) updateUI(app.project);
            } catch (e) {
                alert("Operation failed:\n" + e.toString());
            }
        };

        btnRefresh.onClick = function () {
            updateUI(app.project);
        };

        // ── Init ──────────────────────────────────────────────────────────

        updateUI(app.project);

        if (win instanceof Window) {
            win.center();
            win.show();
        } else {
            win.layout.layout(true);
        }

        return win;
    }

    buildUI(thisObj);

})(this);

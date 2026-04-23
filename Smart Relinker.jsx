// Smart Relinker v1.2
// Finds and relinks missing footage in After Effects projects.
// Folder scanning runs asynchronously to avoid freezing AE.

// Named persistent engine: ensures the panel and the scheduleTask
// string-eval callbacks share the same $.global. On macOS AE 2026
// (stricter than Windows) the panel runs in its own engine by default
// and scheduleTask callbacks fire in the main engine, so a panel that
// registers $.global._srScanStep never sees the tick fire.
#targetengine "SmartRelinker"

(function (thisObj) {

    // =========================================================
    // CHARACTER NORMALIZATION
    // =========================================================
    var CHAR_MAP = {
        '\u00C0':'A','\u00C1':'A','\u00C2':'A','\u00C3':'A','\u00C4':'A','\u00C5':'A',
        '\u00E0':'a','\u00E1':'a','\u00E2':'a','\u00E3':'a','\u00E4':'a','\u00E5':'a',
        '\u00C8':'E','\u00C9':'E','\u00CA':'E','\u00CB':'E',
        '\u00E8':'e','\u00E9':'e','\u00EA':'e','\u00EB':'e',
        '\u00CC':'I','\u00CD':'I','\u00CE':'I','\u00CF':'I',
        '\u00EC':'i','\u00ED':'i','\u00EE':'i','\u00EF':'i',
        '\u00D2':'O','\u00D3':'O','\u00D4':'O','\u00D5':'O','\u00D6':'O','\u00D8':'O',
        '\u00F2':'o','\u00F3':'o','\u00F4':'o','\u00F5':'o','\u00F6':'o','\u00F8':'o',
        '\u00D9':'U','\u00DA':'U','\u00DB':'U','\u00DC':'U',
        '\u00F9':'u','\u00FA':'u','\u00FB':'u','\u00FC':'u',
        '\u00C7':'C','\u00E7':'c','\u00D1':'N','\u00F1':'n',
        '\u00DF':'ss','\u00DD':'Y','\u00FD':'y','\u00FF':'y',
        '\u00C6':'AE','\u00E6':'ae','\u0152':'OE','\u0153':'oe',
        ' ':'_','&':'_and_','@':'_at_'
    };

    // Only collect files with extensions AE can use — skip everything else.
    // Cuts scan time on large drives full of non-media files.
    var AE_EXTS = {
        mp4:1,mov:1,mxf:1,avi:1,r3d:1,braw:1,mkv:1,wmv:1,m4v:1,mpg:1,mpeg:1,webm:1,mp2:1,
        png:1,jpg:1,jpeg:1,tif:1,tiff:1,dpx:1,exr:1,hdr:1,tga:1,bmp:1,psd:1,psb:1,
        ai:1,eps:1,pdf:1,gif:1,cin:1,sgi:1,iff:1,dng:1,raw:1,cr2:1,nef:1,arw:1,
        wav:1,aif:1,aiff:1,mp3:1,aac:1,m4a:1,flac:1,
        c4d:1,obj:1,fbx:1,aep:1,prproj:1,mogrt:1,aepx:1
    };

    // If getFiles() on a folder takes longer than this, treat the folder
    // as pathological (cloud sync stub, huge node_modules, etc.) and do
    // NOT recurse into its subfolders. Files directly in that folder are
    // still kept — only its children are skipped.
    var SLOW_FOLDER_THRESHOLD_MS = 5000;

    function getExtension(name) {
        var dot = name.lastIndexOf('.');
        return dot >= 0 ? name.substring(dot + 1).toLowerCase() : '';
    }

    function normalizeFilename(name) {
        var out = '';
        for (var i = 0; i < name.length; i++) {
            var ch = name.charAt(i);
            out += CHAR_MAP.hasOwnProperty(ch) ? CHAR_MAP[ch] : ch;
        }
        return out.toLowerCase();
    }

    // =========================================================
    // PATH / FILE UTILITIES
    // =========================================================
    function getFileName(path) {
        return path.replace(/\\/g, '/').split('/').pop();
    }

    function comparePathsFromTail(path1, path2, minLevels) {
        var p1 = path1.replace(/\\/g, '/').split('/');
        var p2 = path2.replace(/\\/g, '/').split('/');
        var count = 0;
        var len = Math.min(p1.length, p2.length);
        for (var i = 1; i <= len; i++) {
            if (normalizeFilename(p1[p1.length - i]) === normalizeFilename(p2[p2.length - i])) {
                count++;
            } else {
                break;
            }
        }
        return count >= minLevels;
    }

    function isSequenceItem(item) {
        return /\[\d+-\d+\]/.test(item.name);
    }

    // =========================================================
    // MISSING ITEM DETECTION
    // =========================================================
    function findMissingFootage() {
        var missing = [];
        if (!app.project) return missing;
        for (var i = 1; i <= app.project.numItems; i++) {
            var item = app.project.item(i);
            if (item instanceof FootageItem) {
                try {
                    if (item.mainSource instanceof SolidSource) continue;
                } catch (e) {}
                try {
                    if (item.footageMissing) missing.push(item);
                } catch (e) {}
            }
        }
        return missing;
    }

    function getFontsInProject() {
        var seen = {};
        var fonts = [];
        if (!app.project) return fonts;
        for (var i = 1; i <= app.project.numItems; i++) {
            var item = app.project.item(i);
            if (item instanceof CompItem) {
                for (var j = 1; j <= item.numLayers; j++) {
                    try {
                        var layer = item.layer(j);
                        if (layer instanceof TextLayer) {
                            var font = layer.sourceText.value.font;
                            if (font && !seen[font]) {
                                seen[font] = true;
                                fonts.push(font);
                            }
                        }
                    } catch (e) {}
                }
            }
        }
        return fonts;
    }

    // =========================================================
    // MATCHING
    // =========================================================
    // fileMap: normalizeFilename(file.name) → [File, ...]
    // O(1) for exact matches; only sequence prefix lookups iterate the map.
    function findMatches(item, fileMap) {
        var offlinePath = '';
        try { offlinePath = item.file ? item.file.fsName : ''; } catch (e) {}

        var offlineName = normalizeFilename(getFileName(offlinePath || item.name));
        var isSeq = isSequenceItem(item);
        var seqBase = isSeq ? offlineName.replace(/_?\d+(\.[^.]+)$/, '') : '';

        var candidates = [];

        if (isSeq) {
            for (var key in fileMap) {
                if (key.indexOf(seqBase) === 0) {
                    var arr = fileMap[key];
                    for (var k = 0; k < arr.length; k++) candidates.push(arr[k]);
                }
            }
        } else {
            candidates = fileMap[offlineName] || [];
        }

        var pathMatches = [];
        var nameMatches = [];
        for (var i = 0; i < candidates.length; i++) {
            var f = candidates[i];
            if (offlinePath && comparePathsFromTail(offlinePath, f.fsName, 2)) {
                pathMatches.push(f);
            } else {
                nameMatches.push(f);
            }
        }
        return pathMatches.concat(nameMatches);
    }

    function buildFileMap(files) {
        var map = {};
        for (var i = 0; i < files.length; i++) {
            var key = normalizeFilename(files[i].name);
            if (!map[key]) map[key] = [];
            map[key].push(files[i]);
        }
        return map;
    }

    // =========================================================
    // SCAN LOG — visibility into what's happening during the scan.
    // ScriptUI text can't repaint during scheduleTask, so we write
    // live events to a file the user can tail in a text editor.
    // =========================================================
    // Mac: Desktop (Folder.temp is a per-app sandbox unreadable outside AE).
    // Windows: %TEMP% (standard scratch location, keeps Desktop clean).
    function getLogFile() {
        var dir = ($.os.indexOf('Mac') !== -1) ? Folder.desktop.fsName : Folder.temp.fsName;
        return new File(dir + '/smart_relinker_scan.log');
    }

    function pad2(n) { return n < 10 ? '0' + n : '' + n; }

    function nowStamp() {
        var d = new Date();
        return pad2(d.getHours()) + ':' + pad2(d.getMinutes()) + ':' + pad2(d.getSeconds()) +
               '.' + (d.getMilliseconds() < 100 ? (d.getMilliseconds() < 10 ? '00' : '0') : '') + d.getMilliseconds();
    }

    // Buffered log: we rewrite the whole file on every flush using 'w' mode.
    // macOS ExtendScript's append mode ('a') silently fails in some contexts
    // (including inside scheduleTask callbacks), so we avoid it entirely.
    // The buffer is capped to keep rewrites cheap even during long scans.
    var LOG_BUFFER = [];
    var LOG_MAX = 5000;

    function logFlush() {
        try {
            var f = getLogFile();
            f.encoding = 'UTF-8';
            if (f.open('w')) {
                for (var i = 0; i < LOG_BUFFER.length; i++) {
                    f.writeln(LOG_BUFFER[i]);
                }
                f.close();
            }
        } catch (e) {}
    }

    function logWrite(line) {
        LOG_BUFFER.push('[' + nowStamp() + '] ' + line);
        if (LOG_BUFFER.length > LOG_MAX) LOG_BUFFER.shift();
        logFlush();
    }

    function logReset(header) {
        LOG_BUFFER = [];
        LOG_BUFFER.push('=== Smart Relinker scan log ===');
        LOG_BUFFER.push('Started: ' + new Date().toString());
        if (header) LOG_BUFFER.push(header);
        LOG_BUFFER.push('');
        LOG_BUFFER.push('[' + nowStamp() + '] scan queued — waiting for first scheduleTask tick\u2026');
        LOG_BUFFER.push('(If the next line never appears, the first getFiles() call on the root folder is');
        LOG_BUFFER.push(' blocking. Common cause: Google Drive / SharePoint / network shares with many items.)');
        LOG_BUFFER.push('');
        logFlush();
    }

    function relinkItem(item, file) {
        try {
            if (isSequenceItem(item)) {
                try { item.replaceWithSequence(file, false); return true; } catch (e) {}
            }
            item.replace(file);
            return true;
        } catch (e) { return false; }
    }

    // =========================================================
    // ASYNC SCAN via app.scheduleTask
    // =========================================================
    // One folder per tick keeps AE responsive.
    //
    // Rendering constraints inside a scheduleTask callback:
    //   progressbar.value  → PBM_SETPOS (no WM_PAINT)  — updates instantly ✓
    //   button.enabled     → native state change        — updates instantly ✓
    //   any text property  → requires WM_PAINT          — doesn't update ✗
    //
    // Text labels (folder name, file count) are still set each tick so they
    // paint correctly at end-of-scan when AE's event loop finally runs.
    $.global._srScan = null;
    $.global._srUI   = null;

    $.global._srScanStep = function () {
        logWrite('tick fired');
        var scan = $.global._srScan;
        var ui   = $.global._srUI;

        if (!scan || scan.cancelled) {
            logWrite('CANCELLED after ' + (scan ? scan.processed : 0) + ' folders, ' + (scan ? scan.files.length : 0) + ' files');
            if (ui) {
                try { ui.progBar.value = 0; }                                         catch (e) {}
                try { ui.cancelBtn.enabled = false; }                                 catch (e) {}
                try { ui.scanBtn.text = 'Scan Folder'; ui.scanBtn.enabled = true; }  catch (e) {}
                try { ui.browseBtn.enabled = true; }                                  catch (e) {}
                try { ui.progGroup.visible = false; }                                 catch (e) {}
                try { ui.scanResultLabel.text = 'Scan cancelled.'; }                  catch (e) {}
                try { ui.win.layout.layout(true); }                                   catch (e) {}
                if (ui.onComplete) { try { ui.onComplete([], {}); } catch (e) {} }
            }
            return;
        }

        if (scan.queue.length === 0) {
            scan.done = true;
            var fileMap = buildFileMap(scan.files);
            var slowNote = scan.slowSkipped ? ' (' + scan.slowSkipped + ' slow folder' + (scan.slowSkipped === 1 ? '' : 's') + ' not recursed)' : '';
            logWrite('DONE — ' + scan.processed + ' folders, ' + scan.files.length + ' media files' + slowNote);
            if (ui) {
                try { ui.progBar.value = 100; }                                                                               catch (e) {}
                try { ui.cancelBtn.enabled = false; }                                                                         catch (e) {}
                try { ui.scanBtn.text = scan.files.length + ' files found  \u2014  scan again?'; ui.scanBtn.enabled = true; } catch (e) {}
                try { ui.browseBtn.enabled = true; }                                                                          catch (e) {}
                try { ui.progGroup.visible = false; }                                                                         catch (e) {}
                try { ui.scanResultLabel.text = 'Scanned ' + scan.files.length + ' files.' + (scan.slowSkipped ? ' ' + scan.slowSkipped + ' slow folder(s) skipped — see log.' : ''); } catch (e) {}
                try { ui.win.layout.layout(true); }                                                                           catch (e) {}
            }
            if (ui && ui.onComplete) ui.onComplete(scan.files, fileMap);
            return;
        }

        var folder = scan.queue.shift();

        // Accurate progress: processed / (processed + remaining + current).
        // Denominator can grow when subfolders are discovered, so the bar
        // may dip — that's the honest current state.
        if (ui) {
            var processed = scan.processed;
            var total     = processed + scan.queue.length + 1;
            var pct       = total > 0 ? Math.round((processed / total) * 100) : 0;
            try { ui.progBar.value = pct; } catch (e) {}

            try { ui.scanBtn.text = folder.name + '\u2026'; }                                                catch (e) {}
            try { ui.currentFolderLabel.text = folder.fsName; }                                              catch (e) {}
            try { ui.scanCountLabel.text = scan.files.length + ' files \u00b7 ' + processed + '/' + total + ' folders'; } catch (e) {}
        }

        // Log BEFORE getFiles so a hang is visible in the log.
        logWrite('scan: ' + folder.fsName);
        logWrite('  calling getFiles()\u2026');

        var addedFolders = 0;
        var addedFiles   = 0;
        try {
            var t0 = new Date().getTime();
            var contents = folder.getFiles();
            var dt = new Date().getTime() - t0;
            var slow = dt > SLOW_FOLDER_THRESHOLD_MS;
            logWrite('  getFiles() returned ' + contents.length + ' entries in ' + dt + 'ms' + (slow ? ' [SLOW]' : ''));
            for (var i = 0; i < contents.length; i++) {
                var entry = contents[i];
                if (entry.name.charAt(0) === '.') continue;
                if (entry instanceof Folder) {
                    if (slow) continue;
                    scan.queue.push(entry);
                    addedFolders++;
                } else if (entry instanceof File) {
                    if (AE_EXTS[getExtension(entry.name)]) {
                        scan.files.push(entry);
                        addedFiles++;
                    }
                }
            }
            if (slow) {
                scan.slowSkipped = (scan.slowSkipped || 0) + 1;
                logWrite('  ! SLOW folder — not recursing into its subfolders');
            }
            logWrite('  \u2192 ' + addedFiles + ' media, ' + addedFolders + ' subfolders (queue=' + scan.queue.length + ', total files=' + scan.files.length + ')');
        } catch (e) {
            logWrite('  ! ERROR reading folder: ' + e.toString());
        }

        scan.processed++;

        app.scheduleTask('$.global._srScanStep()', 16, false);
    };

    function startScan(rootFolder, onComplete) {
        $.global._srScan          = { queue: [rootFolder], files: [], processed: 0, done: false, cancelled: false };
        $.global._srUI.onComplete = onComplete;
        logWrite('startScan: scheduling first tick (delay=20ms)');
        try {
            app.scheduleTask('$.global._srScanStep()', 20, false);
        } catch (e) {
            logWrite('  scheduleTask THREW: ' + e.toString());
        }
    }

    function cancelScan() {
        if ($.global._srScan) $.global._srScan.cancelled = true;
    }

    // =========================================================
    // MULTI-MATCH DIALOG
    // =========================================================
    function showMultiMatchDialog(multiItems) {
        var dlg = new Window('dialog', 'Multiple Matches \u2014 Choose the Right File');
        dlg.orientation = 'column';
        dlg.alignChildren = ['fill', 'top'];
        dlg.spacing = 8;
        dlg.margins = 12;

        dlg.add('statictext', undefined,
            multiItems.length + ' item(s) matched more than one file. Pick the correct source for each:');

        var remaining = multiItems.length;

        for (var idx = 0; idx < multiItems.length; idx++) {
            (function (entry) {
                var fp = dlg.add('panel', undefined, '');
                fp.orientation = 'column';
                fp.alignment = ['fill', 'top'];
                fp.alignChildren = ['fill', 'top'];
                fp.spacing = 4;
                fp.margins = 8;

                var offlinePath = '';
                try { offlinePath = entry.item.file ? entry.item.file.fsName : entry.item.name; } catch (e) { offlinePath = entry.item.name; }

                var lbl = fp.add('statictext', undefined, 'Missing: ' + offlinePath);
                lbl.helpTip = offlinePath;

                var radios = [];
                for (var m = 0; m < entry.matches.length; m++) {
                    var rb = fp.add('radiobutton', undefined, entry.matches[m].fsName);
                    rb.helpTip = entry.matches[m].fsName;
                    radios.push(rb);
                }
                if (radios.length > 0) radios[0].value = true;

                var relinkBtn = fp.add('button', undefined, 'Relink This Item');
                relinkBtn.onClick = function () {
                    for (var r = 0; r < radios.length; r++) {
                        if (radios[r].value) {
                            app.beginUndoGroup('Relink ' + entry.item.name);
                            relinkItem(entry.item, entry.matches[r]);
                            app.endUndoGroup();
                            break;
                        }
                    }
                    fp.visible = false;
                    dlg.layout.layout(true);
                    remaining--;
                    if (remaining === 0) dlg.close();
                };
            })(multiItems[idx]);
        }

        dlg.add('button', undefined, 'Close').onClick = function () { dlg.close(); };
        dlg.show();
    }

    // =========================================================
    // UI
    // =========================================================
    function buildUI(thisObj) {
        var isPanel = thisObj instanceof Panel;
        var win = isPanel
            ? thisObj
            : new Window('palette', 'Smart Relinker', undefined, { resizeable: true });

        win.orientation = 'column';
        win.alignChildren = ['fill', 'top'];
        win.spacing = 8;
        win.margins = 10;

        // ---- State ----
        var scannedFiles   = [];
        var scannedFileMap = {};
        var missingItems   = [];
        var isScanning     = false;

        // ---- Header ----
        var hdr = win.add('group');
        hdr.orientation = 'row';
        hdr.alignment = ['fill', 'top'];
        var titleTxt = hdr.add('statictext', undefined, 'Smart Relinker  v1.2');
        titleTxt.graphics.font = ScriptUI.newFont('dialog', 'BOLD', 12);

        // ---- Tabs ----
        var tabs = win.add('tabbedpanel');
        tabs.alignment = ['fill', 'fill'];
        tabs.alignChildren = ['fill', 'fill'];

        var relinkTab = tabs.add('tab', undefined, 'Relink');
        relinkTab.orientation = 'column';
        relinkTab.alignChildren = ['fill', 'top'];
        relinkTab.spacing = 8;
        relinkTab.margins = 8;

        var logTab = tabs.add('tab', undefined, 'Scan Log');
        logTab.orientation = 'column';
        logTab.alignChildren = ['fill', 'fill'];
        logTab.spacing = 6;
        logTab.margins = 8;

        tabs.selection = relinkTab;

        // ---- Missing Items panel ----
        var missingPanel = relinkTab.add('panel', undefined, 'Missing Items');
        missingPanel.orientation = 'column';
        missingPanel.alignChildren = ['fill', 'top'];
        missingPanel.spacing = 6;
        missingPanel.margins = [8, 14, 8, 8];

        var countRow = missingPanel.add('group');
        countRow.orientation = 'row';
        countRow.alignment = ['fill', 'top'];
        var footageCount = countRow.add('statictext', undefined, 'Footage: \u2014');
        var fontCount    = countRow.add('statictext', undefined, '  |  Fonts in project: \u2014');

        var missingList = missingPanel.add('listbox', undefined, [], {
            multiselect: true,
            numberOfColumns: 2,
            columnTitles: ['Filename', 'Original Path'],
            showHeaders: true
        });
        missingList.alignment = ['fill', 'top'];
        // Give the listbox a fixed preferred width so long paths don't force
        // the whole panel wider than the docked window. It still fills
        // horizontally via alignment: ['fill', ...], but the natural width
        // used during layout(true) won't explode.
        missingList.preferredSize = [300, 130];
        missingList.maximumSize = [10000, 130];

        var refreshBtn = missingPanel.add('button', undefined, 'Refresh Missing Items');
        refreshBtn.alignment = ['fill', 'top'];

        // ---- Search Folder panel ----
        var searchPanel = relinkTab.add('panel', undefined, 'Search Folder');
        searchPanel.orientation = 'column';
        searchPanel.alignChildren = ['fill', 'top'];
        searchPanel.spacing = 6;
        searchPanel.margins = [8, 14, 8, 8];

        var folderRow = searchPanel.add('group');
        folderRow.orientation = 'row';
        folderRow.alignment = ['fill', 'top'];
        var folderEdit = folderRow.add('edittext', undefined, 'No folder selected');
        folderEdit.alignment = ['fill', 'center'];
        folderEdit.enabled = false;
        var browseBtn = folderRow.add('button', undefined, 'Browse\u2026');
        browseBtn.preferredSize = [70, -1];

        var scanBtn = searchPanel.add('button', undefined, 'Scan Folder');
        scanBtn.alignment = ['fill', 'top'];
        scanBtn.enabled = false;

        // Progress group — shown only while scanning
        var progGroup = searchPanel.add('group');
        progGroup.orientation = 'column';
        progGroup.alignment = ['fill', 'top'];
        progGroup.alignChildren = ['fill', 'top'];
        progGroup.spacing = 4;
        progGroup.visible = false;

        var progBar = progGroup.add('progressbar', undefined, 0, 100);
        progBar.alignment = ['fill', 'top'];
        progBar.preferredSize = [-1, 24];

        var currentFolderLabel = progGroup.add('statictext', undefined, '', { truncate: 'middle' });
        currentFolderLabel.alignment = ['fill', 'top'];

        var scanCountLabel = progGroup.add('statictext', undefined, '');
        scanCountLabel.alignment = ['fill', 'top'];
        scanCountLabel.justify = 'center';

        var progBtnRow = progGroup.add('group');
        progBtnRow.orientation = 'row';
        progBtnRow.alignment = ['center', 'top'];
        progBtnRow.spacing = 6;

        var cancelBtn = progBtnRow.add('button', undefined, 'Cancel Scan');
        cancelBtn.preferredSize = [110, -1];

        var openLogBtn = progBtnRow.add('button', undefined, 'Open Scan Log');
        openLogBtn.preferredSize = [130, -1];
        openLogBtn.helpTip = 'Opens the live scan log in your default text editor. Tail it to see which folder is being scanned.';

        var scanResultLabel = searchPanel.add('statictext', undefined, 'No folder scanned yet.');
        scanResultLabel.alignment = ['fill', 'top'];
        scanResultLabel.justify = 'center';

        var selectedFolder = null;

        $.global._srUI = {
            win: win,
            progBar: progBar,
            progGroup: progGroup,
            scanBtn: scanBtn,
            browseBtn: browseBtn,
            cancelBtn: cancelBtn,
            currentFolderLabel: currentFolderLabel,
            scanCountLabel: scanCountLabel,
            scanResultLabel: scanResultLabel,
            onComplete: null
        };

        // ---- Relink buttons ----
        var relinkRow = relinkTab.add('group');
        relinkRow.orientation = 'row';
        relinkRow.alignment = ['fill', 'top'];
        var relinkAllBtn = relinkRow.add('button', undefined, 'Relink All');
        relinkAllBtn.alignment = ['fill', 'center'];
        relinkAllBtn.enabled = false;
        var relinkSelBtn = relinkRow.add('button', undefined, 'Relink Selected');
        relinkSelBtn.alignment = ['fill', 'center'];
        relinkSelBtn.enabled = false;

        // ---- Scan Log tab contents ----
        // Not `readonly` — some AE builds suppress programmatic .text updates
        // on readonly edittext. It's effectively read-only since any edits
        // are overwritten on the next Refresh/scan tick.
        var logEdit = logTab.add('edittext', undefined, '(no scan has run yet)', { multiline: true, scrolling: true });
        logEdit.alignment = ['fill', 'fill'];
        logEdit.preferredSize = [-1, 260];
        try {
            var mono = ScriptUI.newFont('Consolas', 'REGULAR', 11) || ScriptUI.newFont('Courier New', 'REGULAR', 11);
            if (mono) logEdit.graphics.font = mono;
        } catch (e) {}

        var logBtnRow = logTab.add('group');
        logBtnRow.orientation = 'row';
        logBtnRow.alignment = ['fill', 'bottom'];
        logBtnRow.spacing = 6;
        var logRefreshBtn = logBtnRow.add('button', undefined, 'Refresh');
        logRefreshBtn.preferredSize = [90, -1];
        var logOpenExtBtn = logBtnRow.add('button', undefined, 'Open Externally');
        logOpenExtBtn.preferredSize = [130, -1];
        logOpenExtBtn.helpTip = 'Opens the log in your default text editor for live tailing (VSCode auto-reloads).';
        var logSpacer = logBtnRow.add('group');
        logSpacer.alignment = ['fill', 'fill'];
        var logPathLabel = logBtnRow.add('statictext', undefined, '', { truncate: 'middle' });
        logPathLabel.alignment = ['fill', 'center'];
        logPathLabel.text = getLogFile().fsName;
        logPathLabel.helpTip = getLogFile().fsName;

        // Show only the tail of large logs — edittext chokes on megabytes,
        // and the tail is the interesting part when hunting hangs.
        var LOG_TAIL_BYTES = 100000;

        function loadLogIntoPanel() {
            var f = getLogFile();
            if (!f.exists) { logEdit.text = '(no scan log yet \u2014 run a scan first)'; return; }
            var opened = false;
            try {
                f.encoding = 'UTF-8';
                if (!f.open('r')) { logEdit.text = '(could not open log \u2014 may be locked, try Refresh)'; return; }
                opened = true;
                var contents;
                if (f.length > LOG_TAIL_BYTES) {
                    f.seek(f.length - LOG_TAIL_BYTES, 0);
                    contents = '\u2026(earlier entries trimmed)\u2026\n' + f.read();
                } else {
                    contents = f.read();
                }
                logEdit.text = contents || '(log is empty)';
            } catch (e) {
                logEdit.text = 'Error reading log: ' + e.toString();
            } finally {
                if (opened) { try { f.close(); } catch (e2) {} }
            }
        }

        logRefreshBtn.onClick = function () { loadLogIntoPanel(); };
        logOpenExtBtn.onClick = function () {
            var f = getLogFile();
            if (!f.exists) {
                try { f.encoding = 'UTF-8'; if (f.open('w')) { f.writeln('(no scan has run yet)'); f.close(); } } catch (e) {}
            }
            try { f.execute(); } catch (e) { alert('Could not open log:\n' + f.fsName); }
        };

        tabs.onChange = function () {
            if (tabs.selection === logTab) loadLogIntoPanel();
        };

        // =================================================
        // Internal helpers
        // =================================================
        function updateRelinkButtons() {
            var hasFiles   = scannedFiles.length > 0;
            var hasMissing = missingItems.length > 0;
            relinkAllBtn.enabled = hasFiles && hasMissing && !isScanning;
            relinkSelBtn.enabled = hasFiles && !isScanning;
        }

        function refreshMissing() {
            missingItems = findMissingFootage();
            var fonts     = getFontsInProject();

            footageCount.text = 'Footage: ' + missingItems.length;
            fontCount.text    = '  |  Fonts in project: ' + fonts.length;

            missingList.removeAll();
            for (var i = 0; i < missingItems.length; i++) {
                var item = missingItems[i];
                var fullPath = '';
                try { fullPath = item.file ? item.file.fsName : item.name; } catch (e) { fullPath = item.name; }
                var row = missingList.add('item', getFileName(fullPath));
                row.subItems[0].text = fullPath;
            }

            updateRelinkButtons();
            // resize() respects the current window bounds; layout(true) would
            // recompute preferred sizes and blow the panel wider than the dock.
            try { win.layout.resize(); } catch (e) {}
            // Move focus off the Refresh button so it doesn't keep the blue ring.
            try { missingList.active = true; } catch (e) {}
        }

        function onScanComplete(files, fileMap) {
            isScanning     = false;
            scannedFiles   = files;
            scannedFileMap = fileMap || {};
            updateRelinkButtons();
            loadLogIntoPanel();
        }

        // =================================================
        // Relink logic
        // =================================================
        function doRelink(selectedOnly) {
            if (scannedFiles.length === 0) { alert('Please scan a folder first.'); return; }

            missingItems = findMissingFootage();

            var toProcess;
            if (selectedOnly) {
                var sel = missingList.selection;
                if (!sel || sel.length === 0) { alert('Select one or more items from the Missing list first.'); return; }
                toProcess = [];
                for (var s = 0; s < sel.length; s++) {
                    if (missingItems[sel[s].index]) toProcess.push(missingItems[sel[s].index]);
                }
            } else {
                toProcess = missingItems.slice();
            }

            if (toProcess.length === 0) { alert('No missing footage to relink.'); return; }

            var autoCount  = 0;
            var multiItems = [];
            var noMatch    = [];

            for (var i = 0; i < toProcess.length; i++) {
                var item    = toProcess[i];
                var matches = findMatches(item, scannedFileMap);

                if (matches.length === 0) {
                    noMatch.push(item);
                } else if (matches.length === 1) {
                    app.beginUndoGroup('Relink ' + item.name);
                    if (relinkItem(item, matches[0])) autoCount++;
                    app.endUndoGroup();
                } else {
                    multiItems.push({ item: item, matches: matches });
                }
            }

            if (multiItems.length > 0) {
                showMultiMatchDialog(multiItems);
            }

            var msg = 'Auto-relinked: ' + autoCount + ' item(s).';
            if (noMatch.length > 0) {
                msg += '\n\nNo match found for ' + noMatch.length + ' item(s):';
                for (var n = 0; n < noMatch.length; n++) {
                    msg += '\n  \u2022 ' + noMatch[n].name;
                }
            }
            if (autoCount > 0 || noMatch.length > 0) alert(msg);

            refreshMissing();
        }

        // =================================================
        // Event wiring
        // =================================================
        refreshBtn.onClick = function () { refreshMissing(); };

        browseBtn.onClick = function () {
            var folder = Folder.selectDialog('Choose a folder to search for footage');
            if (!folder) return;
            selectedFolder = folder;
            folderEdit.text = folder.fsName;
            scannedFiles   = [];
            scannedFileMap = {};
            scanResultLabel.text = 'Folder selected. Click Scan to index it.';
            scanBtn.enabled = true;
            updateRelinkButtons();
            try { win.update(); } catch (e) {}
        };

        scanBtn.onClick = function () {
            if (!selectedFolder) { alert('Please browse for a folder first.'); return; }
            // Take focus to the listbox BEFORE progGroup appears, so the
            // cancel button doesn't briefly own focus and flash its ring.
            try { missingList.active = true; } catch (e) {}

            logReset('Root: ' + selectedFolder.fsName);

            isScanning     = true;
            scannedFiles   = [];
            scannedFileMap = {};
            progBar.value  = 0;
            currentFolderLabel.text = selectedFolder.fsName;
            scanCountLabel.text     = '0 files found';
            progGroup.visible       = true;
            cancelBtn.enabled       = true;
            scanBtn.text            = 'Scanning\u2026';
            scanBtn.enabled         = false;
            browseBtn.enabled       = false;
            scanResultLabel.text    = 'Log: ' + getLogFile().fsName;
            updateRelinkButtons();
            try { win.layout.layout(true); } catch (e) {}
            try { missingList.active = true; } catch (e) {}
            // win.update() throws on macOS Panel objects, killing the handler.
            // layout(true) above is enough — don't call update().
            try { win.update(); } catch (e) {}

            startScan(selectedFolder, function (files, fileMap) { onScanComplete(files, fileMap); });
        };

        cancelBtn.onClick = function () {
            cancelScan();
        };

        openLogBtn.onClick = function () {
            // Launch externally — during a slow getFiles() call AE is frozen,
            // so the in-panel log tab can't live-update. A text editor
            // (VSCode, Notepad++) auto-reloads the file as it grows.
            var f = getLogFile();
            if (!f.exists) {
                try { f.encoding = 'UTF-8'; if (f.open('w')) { f.writeln('(no scan has run yet)'); f.close(); } } catch (e) {}
            }
            try { f.execute(); } catch (e) { alert('Could not open log:\n' + f.fsName); }
        };

        relinkAllBtn.onClick = function () { doRelink(false); };
        relinkSelBtn.onClick = function () { doRelink(true); };

        // Initial state
        refreshMissing();

        if (!isPanel) {
            win.center();
            win.show();
        } else {
            win.layout.layout(true);
            win.layout.resize();
            win.onResize = function () { win.layout.resize(); };
        }

        return win;
    }

    // Preflight: on macOS, Scripting & Expressions > "Allow Scripts to Write
    // Files and Access Network" is OFF by default. Without it every file I/O
    // throws "Permission denied", the scan log never writes, and the UI
    // appears to hang. Detect and explain up front.
    function checkFilePermission() {
        try {
            var probeDir = ($.os.indexOf('Mac') !== -1) ? Folder.desktop.fsName : Folder.temp.fsName;
            var probe = new File(probeDir + '/.smart_relinker_permcheck');
            probe.encoding = 'UTF-8';
            if (probe.open('w')) {
                probe.writeln('ok');
                probe.close();
                probe.remove();
                return true;
            }
            return false;
        } catch (e) { return false; }
    }

    if (!checkFilePermission()) {
        alert(
            'Smart Relinker cannot write files.\n\n' +
            'Enable:\n' +
            '  After Effects > Settings > Scripting & Expressions >\n' +
            '  "Allow Scripts to Write Files and Access Network"\n\n' +
            '(On Windows: Edit > Preferences > Scripting & Expressions.)\n\n' +
            'Then close and reopen this panel.'
        );
        return;
    }

    buildUI(thisObj);

})(this);

/**
 * Auto‑Tile Tool with Editable Symbol Instances
 *
 * - Measures only the portion of the selected object that overlaps the active artboard
 * - Uses a Symbol definition for fast tiling
 * - Breaks each instance link so you keep full paths
 * - Tiles, centers, and ungroups on a new SRA3/SRA3+ sheet
 */

function main() {
    // 1) Preconditions
    if (app.documents.length === 0) {
        alert("Please open a document first!");
        return;
    }
    var doc = app.activeDocument;
    if (doc.selection.length === 0) {
        alert("Please select at least one object!");
        return;
    }
    var item = doc.selection[0];

    // 2) Compute intersection of visibleBounds and artboardRect
    var vb = item.visibleBounds;                // [left, top, right, bottom]
    var abIndex = doc.artboards.getActiveArtboardIndex();
    var abRect  = doc.artboards[abIndex].artboardRect;

    var left   = Math.max(vb[0], abRect[0]);
    var right  = Math.min(vb[2], abRect[2]);
    var top    = Math.min(vb[1], abRect[1]);
    var bottom = Math.max(vb[3], abRect[3]);

    var intersectW = right - left;
    var intersectH = top   - bottom;
    if (intersectW <= 0 || intersectH <= 0) {
        alert("Your selected object does not overlap the active artboard.");
        return;
    }

    // 3) Convert to millimeters for display
    var ptToMm = 25.4 / 72;
    var clippedWmm = (intersectW * ptToMm).toFixed(2);
    var clippedHmm = (intersectH * ptToMm).toFixed(2);

    // 4) Page sizes & dialog
    var mmToPt = 72 / 25.4;
    var pageSizes = {
        "SRA3":  { width: 320 * mmToPt, height: 450 * mmToPt },
        "SRA3+": { width: 330 * mmToPt, height: 488 * mmToPt }
    };

    var dlg = new Window("dialog", "Auto‑Tile Tool");
    dlg.alignChildren = "left";

    var info = dlg.add("panel", undefined, "Measured on Artboard");
    info.orientation = "column";
    info.alignChildren = "left";
    info.add("statictext", undefined, "Width:  " + clippedWmm + " mm");
    info.add("statictext", undefined, "Height: " + clippedHmm + " mm");

    var pg = dlg.add("panel", undefined, "Page Size");
    pg.orientation = "row";
    var rS3  = pg.add("radiobutton", undefined, "SRA3  (320×450 mm)");
    var rS3p = pg.add("radiobutton", undefined, "SRA3+ (330×488 mm)");
    rS3.value = true;

    var mP = dlg.add("panel", undefined, "Margins (mm)");
    mP.orientation = "row";
    mP.alignChildren = "left";
    mP.add("statictext", undefined, "Left:"); var mrL = mP.add("edittext", undefined, "15"); mrL.characters = 4;
    mP.add("statictext", undefined, "Right:");var mrR = mP.add("edittext", undefined, "15"); mrR.characters = 4;
    mP.add("statictext", undefined, "Top:");  var mrT = mP.add("edittext", undefined, "15"); mrT.characters = 4;
    mP.add("statictext", undefined, "Bottom:");var mrB = mP.add("edittext", undefined, "30"); mrB.characters = 4;

    var gapP = dlg.add("panel", undefined, "Gap between copies (mm)");
    gapP.orientation = "row";
    gapP.add("statictext", undefined, "Gap:"); var gapInput = gapP.add("edittext", undefined, "1"); gapInput.characters = 4;

    var btns = dlg.add("group"); btns.orientation = "row";
    btns.add("button", undefined, "OK", { name: "ok" });
    btns.add("button", undefined, "Cancel", { name: "cancel" });
    if (dlg.show() !== 1) return;

    // 5) Read values
    var pageKey     = rS3.value ? "SRA3" : "SRA3+";
    var pageW       = pageSizes[pageKey].width;
    var pageH       = pageSizes[pageKey].height;
    var marginLeft  = parseFloat(mrL.text) * mmToPt;
    var marginRight = parseFloat(mrR.text) * mmToPt;
    var marginTop   = parseFloat(mrT.text) * mmToPt;
    var marginBottom= parseFloat(mrB.text) * mmToPt;
    var gap         = parseFloat(gapInput.text) * mmToPt;

    var objWidth  = intersectW;
    var objHeight = intersectH;

    var availW = pageW - marginLeft - marginRight;
    var availH = pageH - marginTop - marginBottom;
    var cols   = Math.floor((availW + gap) / (objWidth  + gap));
    var rows   = Math.floor((availH + gap) / (objHeight + gap));
    if (cols < 1 || rows < 1) {
        alert("Object too large to fit! Size: " + clippedWmm + "×" + clippedHmm + " mm");
        return;
    }

    // 6) New document
    var newDoc = app.documents.add(DocumentColorSpace.CMYK, pageW, pageH);
    newDoc.rulerUnits = RulerUnits.Points;
    app.redraw();

    // 7) Create symbol and remove master
    var temp = item.duplicate(newDoc, ElementPlacement.PLACEATBEGINNING);
    var sym  = newDoc.symbols.add(temp);
    temp.remove();

    // 8) Group for instances
    var copiesGroup = newDoc.groupItems.add(); copiesGroup.name = "Tiled Copies";

    // 9) Tile symbol instances and break links
    var count = 0;
    for (var r = 0; r < rows; r++) {
        for (var c = 0; c < cols; c++) {
            var x = marginLeft + c*(objWidth+gap);
            var y = pageH - marginTop - r*(objHeight+gap);
            var inst = newDoc.symbolItems.add(sym);
            inst.position = [x,y];
            inst.move(copiesGroup, ElementPlacement.INSIDE);
            // break link so you get full editable paths
            inst.breakLink();
            count++;
        }
    }

    // 10) Center group
    var totalW = cols*objWidth + (cols-1)*gap;
    var totalH = rows*objHeight + (rows-1)*gap;
    var extraW = (availW-totalW)/2;
    var extraH = (availH-totalH)/2;
    copiesGroup.position = [marginLeft+extraW, pageH-marginTop-extraH];

    // 11) Ungroup and summary
    newDoc.selection = null;
    copiesGroup.selected = true;
    app.executeMenuCommand("ungroup");
    newDoc.activate();
    alert(
      "Object on artboard: " + clippedWmm + "×" + clippedHmm + " mm\n" +
      "Grid: "+rows+"×"+cols+" ("+count+" copies)\n" +
      "Margins: L"+mrL.text+" R"+mrR.text+" T"+mrT.text+" B"+mrB.text+" mm\n" +
      "Gap: " + gapInput.text + " mm"
    );
}

main();



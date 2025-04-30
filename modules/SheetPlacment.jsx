/**
 * Auto-Tile Tool (using UnitValue exclusively)
 *
 * - UI and rulers in millimeters
 * - All measurements expressed as UnitValue
 * - No hand-rolled conversion factors
 */

function main() {
    // 1) Preconditions
    if (!app.documents.length) {
        alert("Please open a document first!");
        return;
    }
    var doc = app.activeDocument;
    // show rulers in mm
    app.preferences.rulerUnits    = RulerUnits.Millimeters;
    doc.rulerUnits                = RulerUnits.Millimeters;

    if (!doc.selection.length) {
        alert("Please select something!");
        return;
    }
    var item = doc.selection[0];

    // 2) Compute overlap on active artboard (all in points initially)
    var vb     = item.visibleBounds;    // [L, T, R, B] in pt
    var ab     = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
    var overlapW_pt = Math.min(vb[2], ab[2]) - Math.max(vb[0], ab[0]);
    var overlapH_pt = Math.min(vb[1], ab[1]) - Math.max(vb[3], ab[3]);
    if (overlapW_pt <= 0 || overlapH_pt <= 0) {
        alert("Your selection does not overlap the active artboard.");
        return;
    }

    // wrap into UnitValue
    var wUV = new UnitValue(overlapW_pt, "pt").as("mm");  // numeric mm
    var hUV = new UnitValue(overlapH_pt, "pt").as("mm");  // numeric mm

    // 3) Page sizes, margins & gap as UnitValue in mm
    var pageSizes = {
        "SRA3" : { w: new UnitValue(320, "mm"), h: new UnitValue(450, "mm") },
        "SRA3+": { w: new UnitValue(330, "mm"), h: new UnitValue(488, "mm") }
    };

    // 4) Build dialog (all text in mm)
    var dlg = new Window("dialog", "Auto-Tile Tool (mm)");
    dlg.alignChildren = "left";

    var info = dlg.add("panel", undefined, "Measured on Artboard");
    info.alignChildren = "left";
    info.add("statictext", undefined, "Width:  " + wUV.toFixed(2) + " mm");
    info.add("statictext", undefined, "Height: " + hUV.toFixed(2) + " mm");

    var pg = dlg.add("panel", undefined, "Page Size");
    pg.orientation = "row";
    var rS3  = pg.add("radiobutton", undefined, "SRA3  (320×450 mm)");
    var rS3p = pg.add("radiobutton", undefined, "SRA3+ (330×488 mm)");
    rS3.value = true;

    var mP = dlg.add("panel", undefined, "Margins (mm)");
    mP.orientation = "row";
    mP.add("statictext", undefined, "L:");
    var mrL = mP.add("edittext", undefined, "5");  mrL.characters = 4;
    mP.add("statictext", undefined, "R:");
    var mrR = mP.add("edittext", undefined, "5");  mrR.characters = 4;
    mP.add("statictext", undefined, "T:");
    var mrT = mP.add("edittext", undefined, "15");  mrT.characters = 4;
    mP.add("statictext", undefined, "B:");
    var mrB = mP.add("edittext", undefined, "15");  mrB.characters = 4;

    var gapP = dlg.add("panel", undefined, "Gap between copies (mm)");
    gapP.orientation = "row";
    gapP.add("statictext", undefined, "Gap:");
    var gapInput = gapP.add("edittext", undefined, "1"); gapInput.characters = 4;

    var btns = dlg.add("group");
    btns.orientation = "row";
    btns.add("button", undefined, "OK",     { name: "ok" });
    btns.add("button", undefined, "Cancel", { name: "cancel" });
    if (dlg.show() !== 1) return;

    // 5) Read values (still in mm) and convert to points
    var pageKey    = rS3.value ? "SRA3" : "SRA3+";
    var pgUV       = pageSizes[pageKey];
    var pageW_pt   = pgUV.w.as("pt");
    var pageH_pt   = pgUV.h.as("pt");

    var mL_pt = new UnitValue(parseFloat(mrL.text), "mm").as("pt");
    var mR_pt = new UnitValue(parseFloat(mrR.text), "mm").as("pt");
    var mT_pt = new UnitValue(parseFloat(mrT.text), "mm").as("pt");
    var mB_pt = new UnitValue(parseFloat(mrB.text), "mm").as("pt");
    var gap_pt = new UnitValue(parseFloat(gapInput.text), "mm").as("pt");

    var w_pt = overlapW_pt, h_pt = overlapH_pt;
    var availW = pageW_pt - mL_pt - mR_pt;
    var availH = pageH_pt - mT_pt - mB_pt;
    var cols   = Math.floor((availW + gap_pt) / (w_pt  + gap_pt));
    var rows   = Math.floor((availH + gap_pt) / (h_pt  + gap_pt));
    if (cols < 1 || rows < 1) {
        alert("Object too large to fit! Size: " + wUV.toFixed(2) + "×" + hUV.toFixed(2) + " mm");
        return;
    }

    // 6) Create new document & set rulers to mm
    var newDoc = app.documents.add(DocumentColorSpace.CMYK, pageW_pt, pageH_pt);
    newDoc.rulerUnits = RulerUnits.Millimeters;
    app.redraw();

    // 7) Build symbol in newDoc
    var temp = item.duplicate(newDoc, ElementPlacement.PLACEATBEGINNING);
    var sym  = newDoc.symbols.add(temp);
    temp.remove();

    // 8) Tile and break links
    var copiesGroup = newDoc.groupItems.add();
    copiesGroup.name = "Tiled Copies";
    var count = 0;

    for (var r = 0; r < rows; r++) {
        for (var c = 0; c < cols; c++) {
            var x = mL_pt + c * (w_pt + gap_pt);
            var y = pageH_pt - mT_pt - r * (h_pt + gap_pt);
            var inst = newDoc.symbolItems.add(sym);
            inst.position = [ x, y ];
            inst.move(copiesGroup, ElementPlacement.INSIDE);
            inst.breakLink();
            count++;
        }
    }

    // 9) Center group
    var totalW = cols * w_pt + (cols - 1) * gap_pt;
    var totalH = rows * h_pt + (rows - 1) * gap_pt;
    copiesGroup.position = [
      mL_pt + (availW - totalW) / 2,
      pageH_pt - mT_pt - (availH - totalH) / 2
    ];

    // 10) Ungroup & summary
    newDoc.selection = null;
    copiesGroup.selected = true;
    app.executeMenuCommand("ungroup");
    newDoc.activate();

    alert(
      "Object on artboard: " + wUV.toFixed(2) + "×" + hUV.toFixed(2) + " mm\n" +
      "Grid: " + rows + "×" + cols + " (" + count + " copies)\n" +
      "Margins (mm): L" + mrL.text + " R" + mrR.text +
        " T" + mrT.text + " B" + mrB.text + "\n" +
      "Gap: " + gapInput.text + " mm"
    );

    try {
        app.doScript("Open Print2", "ScriptActions");
    } catch (e) {
        alert(
          "Could not run action “Open Print2” in set “ScriptActions”.\n" +
          "Check that both names match exactly.\n\n" +
          e.message
        );
    }
}

main();

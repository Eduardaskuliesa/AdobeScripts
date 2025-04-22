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
    // just take the first selected item
    var item = doc.selection[0];

    // 2) Get visible bounds of object and artboard
    var vb = item.visibleBounds; // [left, top, right, bottom]
    var abIndex = doc.artboards.getActiveArtboardIndex();
    var abRect  = doc.artboards[abIndex].artboardRect; // [left, top, right, bottom]

    // 3) Compute intersection
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

    // 4) Convert points → millimetres
    var ptToMm = 25.4 / 72;
    var wMM = (intersectW * ptToMm).toFixed(2);
    var hMM = (intersectH * ptToMm).toFixed(2);

    // 5) Alert the result
    alert(
      "Object size on artboard:\n" +
      wMM + " × " + hMM + " mm"
    );
}

main();

// Combined Clone & CutContour Splitter with Auto-Save
(function(){
    // Define constants at the top for easy maintenance
    var SPAUDAI_LOCATION = "F:/PrintReady";
    var PJOVIMUI_LOCATION = "F:/Cut";
    var mmToPt = 2.83465; // mm to points conversion
    
    // Page size definitions
    var pageSizes = {
        "SRA3":  { width: 320 * mmToPt, height: 450 * mmToPt },
        "SRA3+": { width: 330 * mmToPt, height: 488 * mmToPt }
    };
    
    // 0) Preconditions
    if (!app.documents.length) {
        return;
    }
    var origDoc = app.activeDocument,
        ab      = origDoc.artboards[0].artboardRect, // [left, top, right, bottom]
        docW    = ab[2] - ab[0],
        docH    = ab[1] - ab[3];

    // Layers to leave untouched when removing/capturing CutContour
    var protectedLayers = {
      "scpro2_regmarks":     1,
      "scpro2_printonly":    1,
      "scpro2_printmargin":  1
    };
    function isProtectedLayer(name) {
      return !!protectedLayers[name];
    }

    // Fast CutContour detection - exit on first match
    function isCutContour(pi) {
        try {
            // Name check first (fastest)
            if (pi.name === "CutContour") return true;
            
            // Stroke check
            if (pi.stroked && pi.strokeColor.typename === "SpotColor") {
                if (pi.strokeColor.spot.name === "CutContour") return true;
                
                if (pi.strokeColor.spot.color.typename === "RGBColor") {
                    var c = pi.strokeColor.spot.color;
                    if (c.red === 230 && c.green === 46 && c.blue === 146) {
                        return true;
                    }
                }
            }

            // Fill check
            if (pi.filled && pi.fillColor.typename === "SpotColor") {
                if (pi.fillColor.spot.name === "CutContour") return true;
                
                if (pi.fillColor.spot.color.typename === "RGBColor") {
                    var c = pi.fillColor.spot.color;
                    if (c.red === 230 && c.green === 46 && c.blue === 146) {
                        return true;
                    }
                }
            }

            return false;
        } catch(e) {
            return false;
        }
    }

    // 1) Unlock all layers in the original so we can read & duplicate their contents
    for (var L = 0; L < origDoc.layers.length; L++) {
        origDoc.layers[L].locked = false;
    }

    // 2) Clone the document by layer → item duplication
    var cloneDoc = app.documents.add(origDoc.documentColorSpace, docW, docH);
    cloneDoc.rulerUnits = origDoc.rulerUnits;

    // First create all layers to minimize document refresh
    var layerMap = {};
    for (var i = 0; i < origDoc.layers.length; i++) {
        var srcLayer = origDoc.layers[i],
            dstLayer = cloneDoc.layers.add();
        
        // copy layer settings
        dstLayer.name      = srcLayer.name;
        dstLayer.visible   = srcLayer.visible;
        dstLayer.locked    = false;
        dstLayer.template  = srcLayer.template;
        dstLayer.printable = srcLayer.printable;
        
        layerMap[srcLayer.name] = dstLayer;
    }
    
    var cutContourCount = 0;
    var nonCutContourCount = 0;
    
    // Process in batches for better performance
    for (var j = 0; j < origDoc.layers.length; j++) {
        var srcLayer = origDoc.layers[j];
        var dstLayer = layerMap[srcLayer.name];
        
        var itemCount = srcLayer.pageItems.length;
        
        // Skip empty layers
        if (itemCount === 0) continue;
        
        // For large layers, process in batches to show progress
        var batchSize = Math.max(10, Math.min(itemCount, 50));
        var batches = Math.ceil(itemCount / batchSize);
        
        for (var batchIdx = 0; batchIdx < batches; batchIdx++) {
            var startIdx = batchIdx * batchSize;
            var endIdx = Math.min(startIdx + batchSize, itemCount);
            
            // Process this batch
            for (var k = startIdx; k < endIdx; k++) {
                var item = srcLayer.pageItems[k];
                
                try {
                    item.duplicate(dstLayer, ElementPlacement.PLACEATEND);
                } catch(e) {
                    $.writeln("Error duplicating item: " + e);
                }
            }
        }
    }

    // 3) On ORIGINAL: scan ALL pageItems, not just pathItems
    origDoc.activate();
    
    // Directly scan and remove CutContour paths
    var removedCount = 0;
    
    // Recursive function to scan all items including nested ones
    function scanAndRemoveCutContour(parentItem) {
        // If this is an array-like collection of items
        if (parentItem.pageItems && parentItem.pageItems.length > 0) {
            // We need to work with a stable copy of items since the collection will change
            var items = [];
            for (var i = 0; i < parentItem.pageItems.length; i++) {
                items.push(parentItem.pageItems[i]);
            }
            
            // Now process each item
            for (var j = 0; j < items.length; j++) {
                var item = items[j];
                
                // Check if it's a path and a CutContour
                if (item.typename === "PathItem" && isCutContour(item)) {
                    try {
                        item.remove();
                        removedCount++;
                    } catch(e) {
                        $.writeln("Error removing CutContour: " + e);
                    }
                }
                // If it's a container, recursively scan it
                else if (item.typename === "GroupItem" || item.typename === "CompoundPathItem" || 
                         (item.pageItems && item.pageItems.length > 0)) {
                    scanAndRemoveCutContour(item);
                }
            }
        }
    }
    
    // Process by layer
    for (var layerIdx = 0; layerIdx < origDoc.layers.length; layerIdx++) {
        var layer = origDoc.layers[layerIdx];
        var layerName = layer.name;
        
        // Skip protected layers
        if (isProtectedLayer(layerName)) {
            $.writeln("  Skipping protected layer: " + layerName);
            continue;
        }
        
        // Scan this layer recursively
        scanAndRemoveCutContour(layer);
    }

    // OPTIMIZATION FOR CLONE:
    // First ungroup everything in batch mode
    cloneDoc.activate();
    
    function ungroupAll(doc) {
        var groupCount = doc.groupItems.length;
        var originalCount = groupCount;
        var ungrouped = 0;
        
        // Process in batches of 20 for better performance
        while (groupCount > 0) {
            var batchSize = Math.min(20, groupCount);
            var processed = 0;
            
            for (var g = 0; g < batchSize; g++) {
                try {
                    // Always process the first group (index 0) as the collection changes
                    var grp = doc.groupItems[0];
                    var itemCount = grp.pageItems.length;
                    
                    // Move all items out
                    for (var x = itemCount - 1; x >= 0; x--) {
                        try {
                            grp.pageItems[x].move(grp, ElementPlacement.PLACEBEFORE);
                        } catch(e) {
                            // Just continue if we can't move an item
                        }
                    }
                    
                    // Remove the now-empty group
                    grp.remove();
                    processed++;
                    ungrouped++;
                } catch(e) {
                    $.writeln("Error ungrouping: " + e);
                    break;
                }
            }
            
            // Check if we're making progress
            if (processed === 0) {
                $.writeln("Could not process any more groups, stopping ungroup operation");
                break;
            }
            
            // Update count for next iteration
            groupCount = doc.groupItems.length;
        }
        
        return ungrouped;
    }
    
    var ungroupedCount = ungroupAll(cloneDoc);
    
    // Now index all non-CutContour items for faster removal
    var nonCutContourItems = [];
    
    // Scan items in clone doc by layer for better organization
    for (var cloneLayerIdx = 0; cloneLayerIdx < cloneDoc.layers.length; cloneLayerIdx++) {
        var cloneLayer = cloneDoc.layers[cloneLayerIdx];
        var cloneLayerName = cloneLayer.name;
        
        // Skip protected layers
        if (isProtectedLayer(cloneLayerName)) {
            $.writeln("  Skipping protected layer: " + cloneLayerName);
            continue;
        }
        
        var itemCount = cloneLayer.pageItems.length;
        
        if (itemCount > 0) {
            // Process in batches for large layers
            var batchSize = Math.min(1000, Math.max(100, Math.floor(itemCount / 10)));
            
            for (var itemIdx = 0; itemIdx < itemCount; itemIdx++) {
                var item = cloneLayer.pageItems[itemIdx];
                
                // Skip non-path items or items that don't have CutContour properties
                if (item.typename !== "PathItem") {
                    nonCutContourItems.push(item);
                    continue;
                }
                
                // Check if it's a CutContour
                if (!isCutContour(item)) {
                    nonCutContourItems.push(item);
                }
            }
        }
    }
    
    // Now remove all identified non-CutContour items
    var nonCutContourCount = nonCutContourItems.length;
    var keptCount = cloneDoc.pathItems.length - nonCutContourCount;
    
    // Remove in batches for better performance
    var batchSize = Math.min(1000, Math.max(100, Math.floor(nonCutContourCount / 10)));
    
    for (var removeIdx = 0; removeIdx < nonCutContourCount; removeIdx++) {
        try {
            nonCutContourItems[removeIdx].remove();
        } catch(e) {
            // Just continue if we can't remove an item
        }
    }
    
    // ========= AUTO-SAVE FUNCTIONALITY =========
    // This runs after all the CutContour processing is complete
    
    function autoSaveDocuments() {
        try {
            // Find the first non-Untitled document to extract naming information
            var namedDoc = null;
            
            for (var i = 0; i < app.documents.length; i++) {
                var doc = app.documents[i];
                var docName = doc.name;
                
                if (docName.indexOf("Untitled") !== 0) {
                    namedDoc = doc;
                    break;
                }
            }
            
            if (!namedDoc) {
                return null;
            }
            
            // Extract information from the named document's filename
            var fileName = namedDoc.name;
            var parts = fileName.split("_");
            
            if (parts.length < 2) {
                return null;
            }
            
            // Extract order ID and item count
            var orderID = parts[0];
            var itemCount = parts[1];
            var otherInfo = parts.slice(2).join("_");
            
            // Remove file extension if it exists
            if (otherInfo.lastIndexOf('.') != -1) {
                otherInfo = otherInfo.substring(0, otherInfo.lastIndexOf('.'));
            }
            
            // Determine document size
            var docPageSize = "Custom";
            for (var sizeKey in pageSizes) {
                var size = pageSizes[sizeKey];
                // Check with a small tolerance (1pt)
                if (Math.abs(docW - size.width) < 1 && Math.abs(docH - size.height) < 1) {
                    docPageSize = sizeKey;
                    break;
                }
            }
            
            // Calculate papers needed based on item count
            var papersNeeded = 1;
            if (itemCount && !isNaN(parseInt(itemCount, 10))) {
                var numItems = parseInt(itemCount, 10);
                papersNeeded = Math.ceil(numItems / 100);
            }
            
            // Build the base filename
            var baseFilename = orderID + "_" + itemCount + "_" + otherInfo + "_" + docPageSize + "_" + papersNeeded;
            
            // Create folders if they don't exist
            var printReadyFolder = new Folder(SPAUDAI_LOCATION);
           
            if (!printReadyFolder.exists) {
                printReadyFolder.create();
            }
            
            // Set up PDF save options
            var pdfSaveOpts = new PDFSaveOptions();
            pdfSaveOpts.compatibility = PDFCompatibility.ACROBAT8;
            pdfSaveOpts.generateThumbnails = false;
            pdfSaveOpts.preserveEditability = false;
            
            // Save the documents
            origDoc.activate();
            var printFilename = baseFilename + "_spaudai.pdf";
            var printFile = new File(printReadyFolder + "/" + printFilename);
            origDoc.saveAs(printFile, pdfSaveOpts);
            
            origDoc.activate();
            
            return {
                printFile: printFilename,
                printPath: printReadyFolder.fsName,
                papersNeeded: papersNeeded
            };
        }
        catch(e) {
            $.writeln("Error in autoSaveDocuments: " + e);
            return null;
        }
    }
    
    // Run auto-save after all processing is complete
    var saveResults = autoSaveDocuments();

    // try {
    //     app.doScript("Open Print2", "ScriptActions");
    // } catch (e) {
    //     alert(
    //       "Could not run action "Open Print2" in set "ScriptActions".\n" +
    //       "Check that both names match exactly.\n\n" +
    //       e.message
    //     );
    // }
})();
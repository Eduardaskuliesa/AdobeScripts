// Script to preview proposed file operations with paper count in filenames
(function() {
    if (app.documents.length > 0) {
        // Find the first named document
        var namedDoc = null;
        
        for (var i = 0; i < app.documents.length; i++) {
            var doc = app.documents[i];
            var docName = doc.name;
            
            if (docName.indexOf("Untitled") !== 0) {
                namedDoc = doc;
                break;
            }
        }
        
        // Find Untitled-1 and Untitled-2 documents
        var untitled1Doc = null;
        var untitled2Doc = null;
        
        for (var j = 0; j < app.documents.length; j++) {
            var doc = app.documents[j];
            
            if (doc.name === "Untitled-1") {
                untitled1Doc = doc;
            }
            else if (doc.name === "Untitled-2") {
                untitled2Doc = doc;
            }
        }
        
        // Prepare file information
        var message = "PROPOSED FILE OPERATIONS:\n\n";
        
        // Default values in case we can't extract needed info
        var orderID = "ORDERID";
        var itemCount = "COUNT";
        var otherInfo = "additional_info";
        var docPageSize = "Custom";
        
        // Extract info from named document if available
        if (namedDoc) {
            var fileName = namedDoc.name;
            var parts = fileName.split("_");
            
            if (parts.length >= 2) {
                orderID = parts[0];
                itemCount = parts[1];
                otherInfo = parts.slice(2).join("_");
                
                // Remove file extension if it exists
                if (otherInfo.lastIndexOf('.') != -1) {
                    otherInfo = otherInfo.substring(0, otherInfo.lastIndexOf('.'));
                }
            }
        }
        
        // Extract page size from Untitled-2 if available
        if (untitled2Doc) {
            var ab = untitled2Doc.artboards[0].artboardRect;
            var docW = ab[2] - ab[0];
            var docH = ab[1] - ab[3];
            
            var mmToPt = 2.83465;
            var pageSizes = {
                "SRA3":  { width: 320 * mmToPt, height: 450 * mmToPt },
                "SRA3+": { width: 330 * mmToPt, height: 488 * mmToPt }
            };
            
            for (var sizeKey in pageSizes) {
                var size = pageSizes[sizeKey];
                if (Math.abs(docW - size.width) < 1 && Math.abs(docH - size.height) < 1) {
                    docPageSize = sizeKey;
                    break;
                }
            }
        }
        
        // Calculate papers needed based on item count
        var papersNeeded = 1;
        if (itemCount && !isNaN(parseInt(itemCount, 10))) {
            var numItems = parseInt(itemCount, 10);
            papersNeeded = Math.ceil(numItems / 100);
        }
        
        // Create base filename with paper count
        var baseFilename = orderID + "_" + itemCount + "_" + otherInfo + "_" + docPageSize + "_" + papersNeeded;
        
        // Preview Untitled-1 operations
        if (untitled1Doc) {
            message += "DOCUMENT: " + untitled1Doc.name + "\n";
            message += "ACTION: Rename and save to SPAUDAI folder\n";
            message += "NEW NAME: " + baseFilename + "_spaudai.pdf\n";
            message += "LOCATION: F:/PrintReady\n\n";
        } else {
            message += "WARNING: Untitled-1 document not found\n\n";
        }
        
        // Preview Untitled-2 operations
        if (untitled2Doc) {
            message += "DOCUMENT: " + untitled2Doc.name + "\n";
            message += "ACTION: Rename and save to PJOVIMUI folder\n";
            message += "NEW NAME: " + baseFilename + "_pjovimui.pdf\n";
            message += "LOCATION: F:/Cut\n\n";
        } else {
            message += "WARNING: Untitled-2 document not found\n\n";
        }
        
        message += "Papers needed: " + papersNeeded;
        
        alert(message);
    } else {
        alert("No documents are open in Illustrator");
    }
})();
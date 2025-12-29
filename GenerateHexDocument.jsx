/**
 * Generate Hex Document
 *
 * This script creates a new Illustrator document with:
 * - Artboard size: 13cm x 34cm
 * - Three layers: Sponsor (bottom), Hex (middle), Masuri Tab (top)
 * - Artwork copied from the active document
 * - SVG templates imported by opening and copying
 */

// Check if there's an active document
if (app.documents.length === 0) {
    alert("Please open a document before running this script.");
} else {
    main();
}

function main() {
    try {
        // Get the source document (currently active)
        var sourceDoc = app.activeDocument;

        // Get the script's folder path to locate SVG files
        var scriptFile = new File($.fileName);
        var scriptFolder = scriptFile.parent;

        // Define SVG file paths
        var hexSVGFile = new File(scriptFolder + "/HEX.svg");
        var masuriTabSVGFile = new File(scriptFolder + "/MASURI TAB.svg");

        // Verify SVG files exist
        if (!hexSVGFile.exists) {
            alert("HEX.svg not found at: " + hexSVGFile.fsName);
            return;
        }
        if (!masuriTabSVGFile.exists) {
            alert("MASURI TAB.svg not found at: " + masuriTabSVGFile.fsName);
            return;
        }

        // Create new document
        // 13cm = 368.503937 points, 34cm = 963.779528 points
        var docWidth = 368.503937;
        var docHeight = 963.779528;

        var newDoc = app.documents.add(
            DocumentColorSpace.RGB,
            docWidth,
            docHeight
        );

        // Set ruler origin so that the top-left corner of the artboard is at 0,0
        // Get the artboard bounds
        var artboard = newDoc.artboards[0];
        var artboardRect = artboard.artboardRect; // [left, top, right, bottom]

        // Set ruler origin to the top-left corner of the artboard
        newDoc.rulerOrigin = [artboardRect[0], artboardRect[1]];

        // Remove the default layer
        if (newDoc.layers.length > 0) {
            newDoc.layers[0].remove();
        }

        // Create layers in order from bottom to top
        // In Illustrator, the last layer added becomes the topmost layer
        var sponsorLayer = newDoc.layers.add();
        sponsorLayer.name = "Sponsor";

        var hexLayer = newDoc.layers.add();
        hexLayer.name = "Hex";

        var masuriTabLayer = newDoc.layers.add();
        masuriTabLayer.name = "Masuri Tab";

        // Copy all artwork from source document to Sponsor layer
        sourceDoc.activate();
        sourceDoc.selectObjectsOnActiveArtboard();

        if (sourceDoc.selection.length > 0) {
            app.copy();

            newDoc.activate();
            sponsorLayer.locked = false;
            newDoc.activeLayer = sponsorLayer;
            app.paste();

            // Deselect all
            newDoc.selection = null;
        }

        // Import HEX.svg into Hex layer
        newDoc.activate();
        newDoc.activeLayer = hexLayer;
        hexLayer.locked = false;
        importSVGByOpening(hexSVGFile, newDoc, hexLayer);

        // Import MASURI TAB.svg into Masuri Tab layer
        newDoc.activate();
        newDoc.activeLayer = masuriTabLayer;
        masuriTabLayer.locked = false;
        importSVGByOpening(masuriTabSVGFile, newDoc, masuriTabLayer);

        // Group all items on each layer
        groupLayerContents(sponsorLayer);
        groupLayerContents(hexLayer);
        groupLayerContents(masuriTabLayer);

        // Deselect all
        newDoc.selection = null;

        alert("Document created successfully!\n\nLayers (top to bottom):\n1. Masuri Tab\n2. Hex\n3. Sponsor\n\nAll layer contents have been grouped.");

    } catch (e) {
        alert("Error: " + e.message + "\nLine: " + e.line);
    }
}

/**
 * Import SVG file by opening it as a document and copying content
 */
function importSVGByOpening(svgFile, targetDoc, targetLayer) {
    try {
        // Open the SVG file as an Illustrator document
        var svgDoc = app.open(svgFile);

        // Select all content in the SVG
        svgDoc.selectObjectsOnActiveArtboard();

        if (svgDoc.selection.length > 0) {
            // Copy the content
            app.copy();

            // Switch to target document
            targetDoc.activate();
            targetDoc.activeLayer = targetLayer;

            // Paste the content
            app.paste();
        }

        // Close the SVG document without saving
        svgDoc.close(SaveOptions.DONOTSAVECHANGES);

        // Return to target document
        targetDoc.activate();

    } catch (e) {
        throw new Error("Failed to import SVG: " + e.message);
    }
}

/**
 * Group all items on a layer
 */
function groupLayerContents(layer) {
    try {
        var doc = app.activeDocument;

        // Clear selection
        doc.selection = null;

        // Select all page items on this layer
        var items = layer.pageItems;
        if (items.length > 0) {
            // Build selection array
            var selectionArray = [];
            for (var i = 0; i < items.length; i++) {
                selectionArray.push(items[i]);
            }

            // Set the selection
            doc.selection = selectionArray;

            // Group the selection using the built-in group method
            if (doc.selection.length > 0) {
                app.executeMenuCommand("group");
            }
        }

        // Clear selection
        doc.selection = null;

    } catch (e) {
        throw new Error("Failed to group layer contents: " + e.message);
    }
}

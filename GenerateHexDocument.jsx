/**
 * Generate Hex Document
 *
 * This script creates a new Illustrator document with:
 * - Artboard size: 13cm x 34cm
 * - Three layers: Sponsor (bottom), Hex (middle), Masuri Tab (top)
 * - Artwork copied from the active document
 * - SVG templates imported from the repository
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
        var docWidth = 368.503937; // 13cm in points
        var docHeight = 963.779528; // 34cm in points

        var newDoc = app.documents.add(
            DocumentColorSpace.RGB,
            docWidth,
            docHeight
        );

        // Remove the default layer
        if (newDoc.layers.length > 0) {
            newDoc.layers[0].remove();
        }

        // Create layers in reverse order (bottom to top)
        // In Illustrator, layers[0] is the topmost layer
        var masuriTabLayer = newDoc.layers.add();
        masuriTabLayer.name = "Masuri Tab";

        var hexLayer = newDoc.layers.add();
        hexLayer.name = "Hex";

        var sponsorLayer = newDoc.layers.add();
        sponsorLayer.name = "Sponsor";

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
        newDoc.activeLayer = hexLayer;
        hexLayer.locked = false;

        var hexPlaced = newDoc.placedItems.add();
        hexPlaced.file = hexSVGFile;

        // Embed the placed item to preserve vector data
        try {
            hexPlaced.embed();
        } catch (e) {
            // If embed fails, try to open and copy instead
            importSVGAsVectors(hexSVGFile, hexLayer, newDoc);
            hexPlaced.remove();
        }

        // Import MASURI TAB.svg into Masuri Tab layer
        newDoc.activeLayer = masuriTabLayer;
        masuriTabLayer.locked = false;

        var masuriPlaced = newDoc.placedItems.add();
        masuriPlaced.file = masuriTabSVGFile;

        // Embed the placed item to preserve vector data
        try {
            masuriPlaced.embed();
        } catch (e) {
            // If embed fails, try to open and copy instead
            importSVGAsVectors(masuriTabSVGFile, masuriTabLayer, newDoc);
            masuriPlaced.remove();
        }

        // Deselect all
        newDoc.selection = null;

        // Verify layer order (display from top to bottom)
        // layers[0] = Masuri Tab (top)
        // layers[1] = Hex
        // layers[2] = Sponsor (bottom)

        alert("Document created successfully!\n\nLayers (top to bottom):\n1. Masuri Tab\n2. Hex\n3. Sponsor");

    } catch (e) {
        alert("Error: " + e.message + "\nLine: " + e.line);
    }
}

/**
 * Import SVG file as vectors by opening it and copying content
 */
function importSVGAsVectors(svgFile, targetLayer, targetDoc) {
    try {
        // Open the SVG file
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
        throw new Error("Failed to import SVG as vectors: " + e.message);
    }
}

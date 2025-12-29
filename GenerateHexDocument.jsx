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
        // Show color selection dialog
        var hexColor = showColorDialog();
        if (!hexColor) {
            return; // User cancelled
        }

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

            // Scale the pasted sponsor content
            // Target: 11cm wide or 8cm high (whichever is greater in original)
            scaleToFit(newDoc.selection, 11, 8);

            // Deselect all
            newDoc.selection = null;
        }

        // Import HEX.svg into Hex layer
        newDoc.activate();
        newDoc.activeLayer = hexLayer;
        hexLayer.locked = false;
        importSVGByOpening(hexSVGFile, newDoc, hexLayer);

        // Apply selected color to hex layer
        applyColorToLayer(hexLayer, hexColor);

        // Import MASURI TAB.svg into Masuri Tab layer
        newDoc.activate();
        newDoc.activeLayer = masuriTabLayer;
        masuriTabLayer.locked = false;
        importSVGByOpening(masuriTabSVGFile, newDoc, masuriTabLayer);

        // Group all items on each layer
        groupLayerContents(sponsorLayer);
        groupLayerContents(hexLayer);
        groupLayerContents(masuriTabLayer);

        // Position layers
        // Masuri Tab: X = 5.7281cm, Y = 6.8792cm (from top-left reference point)
        positionLayerGroup(masuriTabLayer, 5.7281, 6.8792);

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

/**
 * Position a layer's group to specific coordinates (in cm)
 * Positions the top-left corner of the group's bounding box to the specified coordinates
 * Note: Y coordinate is measured downward from the ruler origin (top-left)
 */
function positionLayerGroup(layer, xCm, yCm) {
    try {
        // Convert cm to points (1 cm = 28.3464567 points)
        var xPoints = xCm * 28.3464567;
        var yPoints = yCm * 28.3464567;

        // Get the first groupItem on the layer (created by groupLayerContents)
        if (layer.groupItems.length > 0) {
            var group = layer.groupItems[0];

            // Get current position (top-left of bounding box)
            var bounds = group.geometricBounds; // [left, top, right, bottom]
            var currentLeft = bounds[0];
            var currentTop = bounds[1];

            // Calculate how much to move
            // In Illustrator, Y increases upward, so moving down is negative
            var deltaX = xPoints - currentLeft;
            var deltaY = -yPoints - currentTop; // Negative because we're moving down from top

            // Translate the group
            group.translate(deltaX, deltaY);
        }
    } catch (e) {
        throw new Error("Failed to position layer group: " + e.message);
    }
}

/**
 * Scale selection to fit within specified dimensions (in cm) while maintaining aspect ratio
 * Scales based on whichever dimension is relatively larger
 */
function scaleToFit(selection, targetWidthCm, targetHeightCm) {
    try {
        if (!selection || selection.length === 0) {
            return;
        }

        // Convert target dimensions from cm to points
        var targetWidthPt = targetWidthCm * 28.3464567;
        var targetHeightPt = targetHeightCm * 28.3464567;

        // Get bounding box of selection
        var bounds = selection[0].geometricBounds;
        for (var i = 1; i < selection.length; i++) {
            var itemBounds = selection[i].geometricBounds;
            bounds[0] = Math.min(bounds[0], itemBounds[0]); // left
            bounds[1] = Math.max(bounds[1], itemBounds[1]); // top
            bounds[2] = Math.max(bounds[2], itemBounds[2]); // right
            bounds[3] = Math.min(bounds[3], itemBounds[3]); // bottom
        }

        // Calculate current dimensions
        var currentWidth = bounds[2] - bounds[0];
        var currentHeight = bounds[1] - bounds[3];

        // Calculate scale factors for each dimension
        var scaleX = (targetWidthPt / currentWidth) * 100; // Convert to percentage
        var scaleY = (targetHeightPt / currentHeight) * 100;

        // Use the smaller scale factor to ensure it fits within the target box
        var scaleFactor = Math.min(scaleX, scaleY);

        // Apply uniform scaling to all selected items
        for (var i = 0; i < selection.length; i++) {
            selection[i].resize(scaleFactor, scaleFactor);
        }

    } catch (e) {
        throw new Error("Failed to scale selection: " + e.message);
    }
}

/**
 * Show dialog to select hex color
 * Returns RGB color object or null if cancelled
 */
function showColorDialog() {
    // Create dialog window
    var dialog = new Window("dialog", "Select Hex Color");
    dialog.orientation = "column";
    dialog.alignChildren = ["fill", "top"];

    // Add dropdown group
    var dropdownGroup = dialog.add("group");
    dropdownGroup.orientation = "row";
    dropdownGroup.add("statictext", undefined, "Hex Color:");

    var colorDropdown = dropdownGroup.add("dropdownlist", undefined, ["Black", "White", "Navy", "Yellow"]);
    colorDropdown.selection = 0; // Default to Black
    colorDropdown.preferredSize.width = 150;

    // Add buttons
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "center";

    var okButton = buttonGroup.add("button", undefined, "OK", {name: "ok"});
    var cancelButton = buttonGroup.add("button", undefined, "Cancel", {name: "cancel"});

    // Show dialog and get result
    if (dialog.show() == 1) {
        // User clicked OK
        var selectedColor = colorDropdown.selection.text;

        // Return RGB color based on selection
        var rgbColor = new RGBColor();

        if (selectedColor == "Black") {
            rgbColor.red = 0;
            rgbColor.green = 0;
            rgbColor.blue = 0;
        } else if (selectedColor == "White") {
            rgbColor.red = 255;
            rgbColor.green = 255;
            rgbColor.blue = 255;
        } else if (selectedColor == "Navy") {
            rgbColor.red = 0;
            rgbColor.green = 0;
            rgbColor.blue = 128;
        } else if (selectedColor == "Yellow") {
            rgbColor.red = 255;
            rgbColor.green = 255;
            rgbColor.blue = 0;
        }

        return rgbColor;
    } else {
        // User cancelled
        return null;
    }
}

/**
 * Apply color to all paths in a layer
 */
function applyColorToLayer(layer, rgbColor) {
    try {
        // Recursively apply color to all path items
        function applyColorToItems(items) {
            for (var i = 0; i < items.length; i++) {
                var item = items[i];

                // If it's a path item, apply the color
                if (item.typename == "PathItem") {
                    item.filled = true;
                    item.fillColor = rgbColor;
                    item.stroked = true;
                    item.strokeColor = rgbColor;
                }
                // If it's a group or compound path, recurse into it
                else if (item.typename == "GroupItem") {
                    applyColorToItems(item.pageItems);
                }
                else if (item.typename == "CompoundPathItem") {
                    applyColorToItems(item.pathItems);
                }
            }
        }

        applyColorToItems(layer.pageItems);

    } catch (e) {
        throw new Error("Failed to apply color to layer: " + e.message);
    }
}

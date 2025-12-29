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
        // Show configuration dialog
        var config = showConfigDialog();
        if (!config) {
            return; // User cancelled
        }

        var hexColor = config.color;
        var sponsorPosition = config.position;

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

        // Position sponsor relative to hex layer
        positionSponsorLayer(sponsorLayer, hexLayer, sponsorPosition);

        // Remove hex paths that overlap with sponsor content
        removeOverlappingHexPaths(hexLayer, sponsorLayer);

        // Import MASURI TAB.svg into Masuri Tab layer
        newDoc.activate();
        newDoc.activeLayer = masuriTabLayer;
        masuriTabLayer.locked = false;
        importSVGByOpening(masuriTabSVGFile, newDoc, masuriTabLayer);

        // Group all items on each layer
        groupLayerContents(sponsorLayer);
        groupLayerContents(hexLayer);
        groupLayerContents(masuriTabLayer);

        // Position layers based on sponsor position mode
        if (sponsorPosition == "Normal Hex Sponsor") {
            // Hex position for Normal mode
            positionLayerGroup(hexLayer, 4.3313, 2.4895);
            // Masuri Tab position for Normal mode
            positionLayerGroup(masuriTabLayer, 5.7281, 6.8792);
        } else if (sponsorPosition == "Sweater Hex Sponsor") {
            // Hex position for Sweater mode
            positionLayerGroup(hexLayer, 4.3313, 2.4895); // TODO: Update with Sweater coordinates
            // Masuri Tab position for Sweater mode
            positionLayerGroup(masuriTabLayer, 5.7281, 6.8792); // TODO: Update with Sweater coordinates
        }

        // Remove any empty layers
        removeEmptyLayers(newDoc);

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
 * Show configuration dialog for hex color and sponsor position
 * Returns object with color and position, or null if cancelled
 */
function showConfigDialog() {
    // Create dialog window
    var dialog = new Window("dialog", "Hex Document Configuration");
    dialog.orientation = "column";
    dialog.alignChildren = ["fill", "top"];

    // Add hex color dropdown group
    var colorGroup = dialog.add("group");
    colorGroup.orientation = "row";
    colorGroup.add("statictext", undefined, "Hex Color:");

    var colorDropdown = colorGroup.add("dropdownlist", undefined, ["Black", "White", "Navy", "Yellow"]);
    colorDropdown.selection = 0; // Default to Black
    colorDropdown.preferredSize.width = 200;

    // Add sponsor position dropdown group
    var positionGroup = dialog.add("group");
    positionGroup.orientation = "row";
    positionGroup.add("statictext", undefined, "Sponsor Position:");

    var positionDropdown = positionGroup.add("dropdownlist", undefined, ["Normal Hex Sponsor", "Sweater Hex Sponsor"]);
    positionDropdown.selection = 0; // Default to Normal
    positionDropdown.preferredSize.width = 200;

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
        var selectedPosition = positionDropdown.selection.text;

        // Create RGB color based on selection
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

        // Return both color and position
        return {
            color: rgbColor,
            position: selectedPosition
        };
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
                    item.stroked = false;
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

/**
 * Remove hex paths that overlap with sponsor layer content
 */
function removeOverlappingHexPaths(hexLayer, sponsorLayer) {
    try {
        // Get bounding box of all sponsor content
        var sponsorBounds = getLayerBounds(sponsorLayer);

        if (!sponsorBounds) {
            return; // No sponsor content, nothing to remove
        }

        // Collect all path items from hex layer (including nested in groups)
        var hexPaths = [];
        collectPathItems(hexLayer.pageItems, hexPaths);

        // Check each hex path for overlap and mark for deletion
        var pathsToDelete = [];
        for (var i = 0; i < hexPaths.length; i++) {
            var pathBounds = hexPaths[i].geometricBounds;

            // Check if bounding boxes intersect
            if (boundsIntersect(pathBounds, sponsorBounds)) {
                pathsToDelete.push(hexPaths[i]);
            }
        }

        // Delete overlapping paths
        for (var i = 0; i < pathsToDelete.length; i++) {
            pathsToDelete[i].remove();
        }

    } catch (e) {
        throw new Error("Failed to remove overlapping hex paths: " + e.message);
    }
}

/**
 * Get the overall bounding box of all items in a layer
 */
function getLayerBounds(layer) {
    var items = layer.pageItems;
    if (items.length === 0) {
        return null;
    }

    var bounds = items[0].geometricBounds;
    for (var i = 1; i < items.length; i++) {
        var itemBounds = items[i].geometricBounds;
        bounds[0] = Math.min(bounds[0], itemBounds[0]); // left
        bounds[1] = Math.max(bounds[1], itemBounds[1]); // top
        bounds[2] = Math.max(bounds[2], itemBounds[2]); // right
        bounds[3] = Math.min(bounds[3], itemBounds[3]); // bottom
    }

    return bounds;
}

/**
 * Recursively collect all path items from a collection
 */
function collectPathItems(items, pathArray) {
    for (var i = 0; i < items.length; i++) {
        var item = items[i];

        if (item.typename == "PathItem") {
            pathArray.push(item);
        } else if (item.typename == "GroupItem") {
            collectPathItems(item.pageItems, pathArray);
        } else if (item.typename == "CompoundPathItem") {
            // Treat compound paths as a single unit (don't break apart)
            pathArray.push(item);
        }
    }
}

/**
 * Check if two bounding boxes intersect
 * Bounds format: [left, top, right, bottom]
 */
function boundsIntersect(bounds1, bounds2) {
    // Two rectangles intersect if:
    // - left1 < right2 AND right1 > left2 (horizontal overlap)
    // - bottom1 < top2 AND top1 > bottom2 (vertical overlap)

    var horizontalOverlap = bounds1[0] < bounds2[2] && bounds1[2] > bounds2[0];
    var verticalOverlap = bounds1[3] < bounds2[1] && bounds1[1] > bounds2[3];

    return horizontalOverlap && verticalOverlap;
}

/**
 * Position sponsor layer relative to hex layer
 * @param {Layer} sponsorLayer - The sponsor layer to position
 * @param {Layer} hexLayer - The hex layer to align with
 * @param {String} positionType - "Normal Hex Sponsor" or "Sweater Hex Sponsor"
 */
function positionSponsorLayer(sponsorLayer, hexLayer, positionType) {
    try {
        // Get bounding boxes
        var sponsorBounds = getLayerBounds(sponsorLayer);
        var hexBounds = getLayerBounds(hexLayer);

        if (!sponsorBounds || !hexBounds) {
            return; // No content to position
        }

        // Calculate current dimensions and centers
        var sponsorWidth = sponsorBounds[2] - sponsorBounds[0];
        var sponsorHeight = sponsorBounds[1] - sponsorBounds[3];
        var sponsorCenterX = sponsorBounds[0] + sponsorWidth / 2;
        var sponsorCenterY = sponsorBounds[3] + sponsorHeight / 2;

        var hexWidth = hexBounds[2] - hexBounds[0];
        var hexHeight = hexBounds[1] - hexBounds[3];
        var hexCenterX = hexBounds[0] + hexWidth / 2;
        var hexCenterY = hexBounds[3] + hexHeight / 2;

        var deltaX = 0;
        var deltaY = 0;

        if (positionType == "Normal Hex Sponsor") {
            // Horizontally centered, bottom edges aligned
            deltaX = hexCenterX - sponsorCenterX;
            deltaY = hexBounds[3] - sponsorBounds[3]; // Align bottom edges
        } else if (positionType == "Sweater Hex Sponsor") {
            // Both horizontally and vertically centered
            deltaX = hexCenterX - sponsorCenterX;
            deltaY = hexCenterY - sponsorCenterY;
        }

        // Move all items in sponsor layer
        var items = sponsorLayer.pageItems;
        for (var i = 0; i < items.length; i++) {
            items[i].translate(deltaX, deltaY);
        }

    } catch (e) {
        throw new Error("Failed to position sponsor layer: " + e.message);
    }
}

/**
 * Remove all empty layers from the document
 */
function removeEmptyLayers(doc) {
    try {
        // Iterate backwards to avoid index shifting issues when removing layers
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            var layer = doc.layers[i];

            // Check if layer is empty (no page items)
            if (layer.pageItems.length === 0) {
                layer.remove();
            }
        }
    } catch (e) {
        // Don't throw error for layer removal - just continue
        // This prevents script failure if layer can't be removed
    }
}

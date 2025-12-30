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

/**
 * Orchestrates creation of a 13cm × 34cm CMYK Illustrator document assembled from source artwork and SVG templates.
 *
 * Creates layers (Sponsor, Hex, Masuri Tab), imports artwork from the active document and the SVG files
 * (HEX.svg, MASURI TAB.svg, GUIDES.svg) located next to the script, applies the selected hex color to the Hex layer,
 * positions and scales sponsor artwork relative to the hex, removes hex paths that overlap sponsor content,
 * converts imported guide artwork to Illustrator guides, groups and centers layer contents, removes empty layers,
 * and displays a success or error alert.
 */
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
        var guidesSVGFile = new File(scriptFolder + "/GUIDES.svg");

        // Verify SVG files exist
        if (!hexSVGFile.exists) {
            alert("HEX.svg not found at: " + hexSVGFile.fsName);
            return;
        }
        if (!masuriTabSVGFile.exists) {
            alert("MASURI TAB.svg not found at: " + masuriTabSVGFile.fsName);
            return;
        }
        if (!guidesSVGFile.exists) {
            alert("GUIDES.svg not found at: " + guidesSVGFile.fsName);
            return;
        }

        // Create new document with centimeter units
        // 13cm x 34cm
        var docPreset = new DocumentPreset();
        docPreset.units = RulerUnits.Centimeters;  // Set units FIRST
        docPreset.width = 368.504;
        docPreset.height = 963.78;
        docPreset.colorMode = DocumentColorSpace.CMYK;

        var newDoc = app.documents.addDocument("Print", docPreset);

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

        // Filter out guides from the selection
        if (sourceDoc.selection.length > 0) {
            var filteredSelection = [];
            for (var i = 0; i < sourceDoc.selection.length; i++) {
                var item = sourceDoc.selection[i];
                // Only include items that are not guides
                if (!item.guides) {
                    filteredSelection.push(item);
                }
            }

            // Update selection to exclude guides
            sourceDoc.selection = filteredSelection;
        }

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

        // Import GUIDES.svg and convert to guides
        var guidesLayer = importGuidesFromSVG(guidesSVGFile, newDoc);

        // Group all items on each layer
        groupLayerContents(sponsorLayer);
        groupLayerContents(hexLayer);
        groupLayerContents(masuriTabLayer);
        if (guidesLayer) {
            groupLayerContents(guidesLayer);
            // Lock the Guides layer
            guidesLayer.locked = true;
        }

        // Name the groups
        if (sponsorLayer.groupItems.length > 0) {
            sponsorLayer.groupItems[0].name = "Sponsor";
        }
        if (hexLayer.groupItems.length > 0) {
            hexLayer.groupItems[0].name = "Hex";
        }
        if (masuriTabLayer.groupItems.length > 0) {
            masuriTabLayer.groupItems[0].name = "Masuri Tab";
        }

        // Position layers based on sponsor position mode
        if (sponsorPosition == "Normal Hex Sponsor") {
            // Hex position for Normal mode
            positionLayerGroup(hexLayer, 4.3313, 2.4895);
            // Masuri Tab position for Normal mode
            positionLayerGroup(masuriTabLayer, 5.7135, 18.8111);
        } else if (sponsorPosition == "Sweater Hex Sponsor") {
            // Hex position for Sweater mode
            positionLayerGroup(hexLayer, 4.3313, 2.4895);
            // Masuri Tab position for Sweater mode
            positionLayerGroup(masuriTabLayer, 5.7135, 8.4606);
        }

        // Group all content layers together and center on artboard
        groupAndCenterLayers(newDoc, [sponsorLayer, hexLayer, masuriTabLayer]);

        // Remove any empty layers
        removeEmptyLayers(newDoc);

        // Deselect all
        newDoc.selection = null;

        alert("Hex created successfully!");

    } catch (e) {
        alert("Error: " + e.message + "\nLine: " + e.line);
    }
}

/**
 * Import an SVG's artwork into a specified layer by opening the SVG, copying its artboard content, and pasting it into the target document.
 * @param {File} svgFile - The SVG file to open and import.
 * @param {Document} targetDoc - The Illustrator document that will receive the pasted artwork.
 * @param {Layer} targetLayer - The layer in the target document where the SVG content will be pasted.
 * @throws {Error} If the SVG cannot be opened, copied, or pasted into the target document.
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
 * Group all page items contained directly on the given layer into a single group.
 * @param {Layer} layer - The Illustrator layer whose top-level pageItems will be grouped.
 * @throws {Error} If grouping fails.
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
 * Aligns the top-left corner of the first group on a layer to the specified coordinates on the artboard.
 * @param {Layer} layer - Illustrator layer whose first groupItem will be positioned.
 * @param {number} xCm - Target x coordinate in centimeters measured from the left ruler origin.
 * @param {number} yCm - Target y coordinate in centimeters measured from the top ruler origin (measured downward).
 * @throws {Error} If positioning fails.
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
 * Scale the given selection to fit within the specified width and height in centimeters while preserving aspect ratio.
 *
 * @param {Array} selection - Array of page items (e.g., PageItem, GroupItem) to be scaled; if empty or null the function does nothing.
 * @param {number} targetWidthCm - Target maximum width in centimeters.
 * @param {number} targetHeightCm - Target maximum height in centimeters.
 * @throws {Error} If an error occurs while computing bounds or applying the resize. 
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
 * Present a dialog to choose the hex color and sponsor position.
 * @returns {{color: RGBColor, position: string} | null} An object with `color` set to an Illustrator `RGBColor` for the selected hex color and `position` set to either "Normal Hex Sponsor" or "Sweater Hex Sponsor", or `null` if the user cancels.
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

        // Create RGB color based on selection using hex codes
        var rgbColor;

        if (selectedColor == "Black") {
            rgbColor = hexToRGB("#000000");
        } else if (selectedColor == "White") {
            rgbColor = hexToRGB("#FFFFFF");
        } else if (selectedColor == "Navy") {
            rgbColor = hexToRGB("#000080");
        } else if (selectedColor == "Yellow") {
            rgbColor = hexToRGB("#FFFF00");
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
 * Apply an Illustrator RGBColor to every path in a layer, traversing nested GroupItem and CompoundPathItem containers.
 * @param {Layer} layer - The layer whose path items will be updated.
 * @param {RGBColor} rgbColor - The color to apply to path fills; affected paths will be filled with this color and stroked will be disabled.
 */
function applyColorToLayer(layer, rgbColor) {
    try {
        /**
         * Recursively applies the external `rgbColor` to every PathItem within the provided collection.
         *
         * Traverses GroupItem and CompoundPathItem structures and for each PathItem sets `filled = true`, assigns `fillColor = rgbColor`, and disables stroke (`stroked = false`).
         * @param {Array} items - A collection of page items (e.g., layer.pageItems, group.pageItems) to process.
         */
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
 * Removes path items on the hex layer whose bounding boxes intersect the combined bounds of the sponsor layer.
 * @param {Layer} hexLayer - The layer containing hex path items to evaluate and remove.
 * @param {Layer} sponsorLayer - The layer whose combined content bounds are used to test for overlap.
 * @throws {Error} If the removal operation fails.
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
 * Compute the combined geometric bounding box for all page items in a layer.
 * @param {Layer} layer - The Illustrator Layer to measure.
 * @returns {(number[]|null)} An array [left, top, right, bottom] in document points representing the bounding box, or `null` if the layer contains no page items.
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
 * Collects PathItem and CompoundPathItem objects from a collection into an array, traversing nested GroupItems.
 * @param {Iterable} items - A collection of page items (e.g., layer.pageItems or group.pageItems) to traverse.
 * @param {Array} pathArray - Array to which found PathItem and CompoundPathItem objects will be appended.
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
 * Determine whether two axis-aligned bounding boxes intersect.
 * 
 * @param {number[]} bounds1 - Bounds as [left, top, right, bottom] for the first rectangle.
 * @param {number[]} bounds2 - Bounds as [left, top, right, bottom] for the second rectangle.
 * @returns {boolean} `true` if the rectangles overlap horizontally and vertically, `false` otherwise.
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
 * Aligns the sponsor layer relative to the hex layer according to the chosen positioning mode.
 *
 * Positions the sponsor layer by translating all its page items so they are either horizontally centered and bottom-aligned with the hex layer ("Normal Hex Sponsor"), or both horizontally and vertically centered within the hex layer ("Sweater Hex Sponsor"). If either layer has no content, the function does nothing.
 *
 * @param {Layer} sponsorLayer - Layer containing sponsor artwork to be moved.
 * @param {Layer} hexLayer - Layer whose bounds are used as the alignment reference.
 * @param {String} positionType - Alignment mode: "Normal Hex Sponsor" or "Sweater Hex Sponsor".
 * @throws {Error} If positioning fails due to an unexpected runtime error.
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
 * Remove layers that contain no page items from the given document.
 * Any errors encountered while removing individual layers are ignored.
 * @param {Document} doc - The Illustrator document to clean of empty layers.
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

/**
 * Create a "Guides" layer in the target document, import all artwork from the given SVG into that layer, translate the pasted items so their top-left sits 0.5 cm from the artboard top-left, and convert the pasted items into Illustrator guides.
 * @param {File} svgFile - The SVG file to import.
 * @param {Document} targetDoc - The Illustrator document to receive the guides.
 * @returns {Layer} The newly created Guides layer containing the converted guide items.
 * @throws {Error} If the SVG cannot be opened, pasted, positioned, or converted to guides.
 */
function importGuidesFromSVG(svgFile, targetDoc) {
    try {
        // Create a Guides layer
        var guidesLayer = targetDoc.layers.add();
        guidesLayer.name = "Guides";

        // Open the SVG file
        var svgDoc = app.open(svgFile);

        // Select all content in the SVG
        svgDoc.selectObjectsOnActiveArtboard();

        if (svgDoc.selection.length > 0) {
            // Copy the content
            app.copy();

            // Switch to target document
            targetDoc.activate();

            // Set Guides layer as active
            targetDoc.activeLayer = guidesLayer;

            // Paste the content
            app.paste();

            // Get the pasted items (they're still selected)
            var pastedItems = targetDoc.selection;

            // Position the guides 0.5cm from top-left corner (same for both modes)
            // Convert 0.5cm to points
            var targetX = 0.5 * 28.3464567; // 0.5cm in points
            var targetY = -0.5 * 28.3464567; // 0.5cm downward (negative in Illustrator coords)

            // Get bounding box of all pasted items
            if (pastedItems.length > 0) {
                var bounds = pastedItems[0].geometricBounds;
                for (var i = 1; i < pastedItems.length; i++) {
                    var itemBounds = pastedItems[i].geometricBounds;
                    bounds[0] = Math.min(bounds[0], itemBounds[0]); // left
                    bounds[1] = Math.max(bounds[1], itemBounds[1]); // top
                    bounds[2] = Math.max(bounds[2], itemBounds[2]); // right
                    bounds[3] = Math.min(bounds[3], itemBounds[3]); // bottom
                }

                // Calculate translation to move top-left corner to target position
                var deltaX = targetX - bounds[0];
                var deltaY = targetY - bounds[1];

                // Apply translation to all items
                for (var i = 0; i < pastedItems.length; i++) {
                    pastedItems[i].translate(deltaX, deltaY);
                }
            }

            // Convert all pasted items to guides
            convertToGuides(pastedItems);
        }

        // Close the SVG document without saving
        svgDoc.close(SaveOptions.DONOTSAVECHANGES);

        // Return to target document
        targetDoc.activate();

        // Return the guides layer
        return guidesLayer;

    } catch (e) {
        throw new Error("Failed to import guides: " + e.message);
    }
}

/**
 * Recursively convert PathItem objects within a collection to guides.
 *
 * Traverses the provided collection of page items and marks each PathItem as a guide.
 * If a GroupItem or CompoundPathItem is encountered, its child items are processed recursively.
 *
 * @param {Array} items - An array or array-like collection of page items (PageItems, GroupItems, CompoundPathItems) to convert.
 */
function convertToGuides(items) {
    for (var i = 0; i < items.length; i++) {
        var item = items[i];

        if (item.typename == "PathItem") {
            item.guides = true;
        } else if (item.typename == "GroupItem") {
            convertToGuides(item.pageItems);
        } else if (item.typename == "CompoundPathItem") {
            convertToGuides(item.pathItems);
        }
    }
}

/**
 * Convert a hex color string into RGB values.
 *
 * @param {string} hex - Hex color string in 3- or 6-digit form, with or without a leading `#` (e.g., `#FFF` or `FFFFFF`).
 * @return {RGBColor} An Illustrator `RGBColor` whose `red`, `green`, and `blue` fields are set to the corresponding 0–255 values.
 */
function hexToRGB(hex) {
    // Remove # if present
    hex = hex.replace("#", "");

    // Handle shorthand hex (e.g., #FFF -> #FFFFFF)
    if (hex.length === 3) {
        hex = hex.charAt(0) + hex.charAt(0) +
              hex.charAt(1) + hex.charAt(1) +
              hex.charAt(2) + hex.charAt(2);
    }

    // Parse hex values
    var r = parseInt(hex.substring(0, 2), 16);
    var g = parseInt(hex.substring(2, 4), 16);
    var b = parseInt(hex.substring(4, 6), 16);

    // Create and return RGBColor
    var rgbColor = new RGBColor();
    rgbColor.red = r;
    rgbColor.green = g;
    rgbColor.blue = b;

    return rgbColor;
}

/**
 * Group the first group item from each provided layer into a single "Artwork" group, rename its layer to "Artwork", and center that group on the document's first artboard.
 * @param {Document} doc - The Illustrator document containing the layers.
 * @param {Array<Layer>} layers - Array of Layer objects whose first groupItems (if present) will be combined.
 * @throws {Error} If grouping or translation fails.
 */
function groupAndCenterLayers(doc, layers) {
    try {
        // Clear selection
        doc.selection = null;

        // Select all groups from the specified layers
        var allGroups = [];
        for (var i = 0; i < layers.length; i++) {
            if (layers[i].groupItems.length > 0) {
                allGroups.push(layers[i].groupItems[0]);
            }
        }

        // If we have groups to work with
        if (allGroups.length > 0) {
            // Set selection to all groups
            doc.selection = allGroups;

            // Group them together
            app.executeMenuCommand("group");

            // Get the master group (the selection after grouping)
            var masterGroup = doc.selection[0];

            // Name the group "Artwork"
            masterGroup.name = "Artwork";

            // Rename the layer containing the artwork to "Artwork"
            if (masterGroup.layer) {
                masterGroup.layer.name = "Artwork";
            }

            // Get artboard bounds
            var artboard = doc.artboards[0];
            var artboardRect = artboard.artboardRect; // [left, top, right, bottom]

            // Calculate artboard center
            var artboardCenterX = (artboardRect[0] + artboardRect[2]) / 2;
            var artboardCenterY = (artboardRect[1] + artboardRect[3]) / 2;

            // Get group bounds
            var groupBounds = masterGroup.geometricBounds; // [left, top, right, bottom]

            // Calculate group center
            var groupCenterX = (groupBounds[0] + groupBounds[2]) / 2;
            var groupCenterY = (groupBounds[1] + groupBounds[3]) / 2;

            // Calculate translation needed to center the group
            var deltaX = artboardCenterX - groupCenterX;
            var deltaY = artboardCenterY - groupCenterY;

            // Move the group to center
            masterGroup.translate(deltaX, deltaY);
        }

        // Clear selection
        doc.selection = null;

    } catch (e) {
        throw new Error("Failed to group and center layers: " + e.message);
    }
}
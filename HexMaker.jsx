/**
 * Generate Hex Document
 *
 * This script creates a new Illustrator document with:
 * - Artboard size: 13cm x 34cm
 * - Three layers: Sponsor (bottom), Hex (middle), Masuri Tab (top)
 * - Artwork copied from the active document
 * - SVG templates imported by opening and copying
 */



// Constants
var POINTS_PER_CM = 28.3464567;

// Check if there's an active document
if (app.documents.length === 0) {
    alert("Please open a document before running this script.");
} else {
    main();
}

/**
 * Builds a new 13cm × 34cm CMYK Illustrator document by importing templates, copying source artwork, and composing the final grouped artwork.
 *
 * Opens a configuration dialog to choose hex color and sponsor position, validates required SVG templates (HEX.svg, MASURI TAB.svg, GUIDES.svg) in the script folder, creates the document with three layers (Sponsor, Hex, Masuri Tab), copies non-guide artwork from the active document into the Sponsor layer and scales it to fit, imports the SVGs into their respective layers (converting one set to guides), applies the chosen color to the Hex layer, aligns and positions layers according to the selected sponsor mode, removes overlapping hex paths, groups and centers content into a final "Artwork" group, removes empty layers, and notifies the user on success. Errors are caught and reported via an alert with the error message and line number.
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

        // Define template and SVG file paths
        var hexTemplateFile = new File(scriptFolder + "/assets/HEX.eps");
        var masuriTabSVGFile = new File(scriptFolder + "/assets/MASURI TAB.svg");
        var guidesSVGFile = new File(scriptFolder + "/assets/GUIDES.svg");

        // Verify template and SVG files exist
        if (!hexTemplateFile.exists) {
            alert("HEX.eps template not found at: " + hexTemplateFile.fsName);
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

        // Open HEX.eps template (already 13cm × 34cm CMYK with hex pattern)
        var newDoc = app.open(hexTemplateFile);

        // OPTION 4 PROTECTION: Clear file path so template can't be overwritten
        // This makes Illustrator treat it as an unsaved document
        // User will be prompted with "Save As" dialog instead of "Save"
        try {
            newDoc.saved = false;
        } catch (e) {
            // If saved property can't be set, continue anyway
        }

        app.preferences.setBooleanPreference("showTransparencyGrid", true);

        // Get artboard bounds for ruler origin
        var artboard = newDoc.artboards[0];
        var artboardRect = artboard.artboardRect;

        // Find the existing layer in template and rename to "Artwork"
        var artworkLayer = null;
        for (var i = 0; i < newDoc.layers.length; i++) {
            if (newDoc.layers[i].name == "Hex" || newDoc.layers[i].name == "Layer 1" || newDoc.layers[i].name == "Artwork") {
                artworkLayer = newDoc.layers[i];
                artworkLayer.name = "Artwork"; // Rename to Artwork
                break;
            }
        }

        // If artwork layer not found, use the first layer
        if (!artworkLayer && newDoc.layers.length > 0) {
            artworkLayer = newDoc.layers[0];
            artworkLayer.name = "Artwork";
        }

        // Copy all artwork from source document to Artwork layer
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

        // Store sponsor content for later grouping
        var sponsorItems = [];
        if (sourceDoc.selection.length > 0) {
            app.copy();

            newDoc.activate();
            artworkLayer.locked = false;
            newDoc.activeLayer = artworkLayer;
            app.paste();

            // Scale the pasted sponsor content
            // Target: 11cm wide or 8cm high (whichever is greater in original)
            scaleToFit(newDoc.selection, 11, 8);

            // Store sponsor items for grouping
            for (var i = 0; i < newDoc.selection.length; i++) {
                sponsorItems.push(newDoc.selection[i]);
            }

            // Deselect all
            newDoc.selection = null;
        }

        // Hex pattern already exists in template on Artwork layer - just apply color
        newDoc.activate();
        newDoc.activeLayer = artworkLayer;
        artworkLayer.locked = false;

        // Store hex items before adding more content
        var hexItems = [];
        for (var i = 0; i < artworkLayer.pageItems.length; i++) {
            var item = artworkLayer.pageItems[i];
            // Skip sponsor items we just added
            var isSponsor = false;
            for (var j = 0; j < sponsorItems.length; j++) {
                if (item == sponsorItems[j]) {
                    isSponsor = true;
                    break;
                }
            }
            if (!isSponsor) {
                hexItems.push(item);
            }
        }

        // Apply selected color to hex items
        applyColorToItems(hexItems, hexColor);

        // Import MASURI TAB.svg into Artwork layer
        newDoc.activate();
        newDoc.activeLayer = artworkLayer;
        artworkLayer.locked = false;

        var itemCountBefore = artworkLayer.pageItems.length;
        importSVGByOpening(masuriTabSVGFile, newDoc, artworkLayer);

        // Store Masuri Tab items
        var masuriTabItems = [];
        for (var i = itemCountBefore; i < artworkLayer.pageItems.length; i++) {
            masuriTabItems.push(artworkLayer.pageItems[i]);
        }

        // Import GUIDES.svg and convert to guides
        var guidesLayer = importGuidesFromSVG(guidesSVGFile, newDoc);
        if (guidesLayer) {
            groupLayerContents(guidesLayer);
            guidesLayer.locked = true;
        }

        // Group each set of items separately
        newDoc.selection = null;

        // Group Sponsor items
        var sponsorGroup = null;
        if (sponsorItems.length > 0) {
            newDoc.selection = sponsorItems;
            app.executeMenuCommand("group");
            if (newDoc.selection.length > 0) {
                sponsorGroup = newDoc.selection[0];
                sponsorGroup.name = "Sponsor";
            }
            newDoc.selection = null;
        }

        // Group Hex items
        var hexGroup = null;
        if (hexItems.length > 0) {
            newDoc.selection = hexItems;
            app.executeMenuCommand("group");
            if (newDoc.selection.length > 0) {
                hexGroup = newDoc.selection[0];
                hexGroup.name = "Hex";
            }
            newDoc.selection = null;
        }

        // Group Masuri Tab items
        var masuriTabGroup = null;
        if (masuriTabItems.length > 0) {
            newDoc.selection = masuriTabItems;
            app.executeMenuCommand("group");
            if (newDoc.selection.length > 0) {
                masuriTabGroup = newDoc.selection[0];
                masuriTabGroup.name = "Masuri Tab";
            }
            newDoc.selection = null;
        }

        // Group all three groups together FIRST, before positioning
        // This avoids issues with the group command failing on moved items
        var masterGroup = null;

        // Ensure all groups are unlocked and visible
        if (sponsorGroup) {
            sponsorGroup.locked = false;
            sponsorGroup.hidden = false;
        }
        if (hexGroup) {
            hexGroup.locked = false;
            hexGroup.hidden = false;
        }
        if (masuriTabGroup) {
            masuriTabGroup.locked = false;
            masuriTabGroup.hidden = false;
        }

        var sponsorHexGroup = null;
        if (sponsorGroup && hexGroup && masuriTabGroup) {
            // WORKAROUND: Can't select all 3 together, so group in two steps
            // Step 1: Group sponsor and hex together
            newDoc.selection = [sponsorGroup, hexGroup];
            if (newDoc.selection.length === 2) {
                app.executeMenuCommand("group");
                if (newDoc.selection.length > 0) {
                    sponsorHexGroup = newDoc.selection[0];
                    sponsorHexGroup.name = "SponsorHex_Temp";
                    sponsorHexGroup.locked = false;
                    sponsorHexGroup.hidden = false;
                }
                newDoc.selection = null;
            }

            // Debug: Check if both groups exist and can be selected
            alert("Before step 2:\nsponsorHexGroup exists: " + (sponsorHexGroup != null) +
                  "\nmasuriTabGroup exists: " + (masuriTabGroup != null) +
                  "\nsponsorHex layer: " + (sponsorHexGroup ? sponsorHexGroup.layer.name : "null") +
                  "\nmasuri layer: " + (masuriTabGroup ? masuriTabGroup.layer.name : "null"));

            // Step 2: Group that with masuri tab to create final Artwork group
            if (sponsorHexGroup && masuriTabGroup) {
                // Try selecting sponsorHexGroup alone
                newDoc.selection = [sponsorHexGroup];
                var sponsorHexSelects = (newDoc.selection.length == 1);

                // Try selecting masuriTabGroup alone
                newDoc.selection = [masuriTabGroup];
                var masuriSelects = (newDoc.selection.length == 1);

                alert("Can select for step 2?\nsponsorHex: " + sponsorHexSelects + "\nmasuri: " + masuriSelects);

                newDoc.selection = [sponsorHexGroup, masuriTabGroup];
                alert("Combined selection length: " + newDoc.selection.length);

                if (newDoc.selection.length === 2) {
                    app.executeMenuCommand("group");
                    if (newDoc.selection.length > 0) {
                        masterGroup = newDoc.selection[0];
                        masterGroup.name = "Artwork";
                    }
                } else {
                    alert("ERROR: Could not select both. Selected: " + newDoc.selection.length);
                }
                newDoc.selection = null;
            }
        }

        // Now position the individual groups within the master group
        // Position sponsor relative to hex
        if (sponsorGroup && hexGroup) {
            positionGroupRelativeToGroup(sponsorGroup, hexGroup, sponsorPosition);
        }

        // Remove hex paths that overlap with sponsor
        if (hexGroup && sponsorGroup) {
            removeOverlappingPathsInGroups(hexGroup, sponsorGroup);
        }

        // Center hex and sponsor together
        if (hexGroup && sponsorGroup) {
            centerTwoGroupsOnArtboard(hexGroup, sponsorGroup, artboardRect);
        }

        // Position Masuri Tab
        if (masuriTabGroup) {
            if (sponsorPosition == "Bottom Sponsor") {
                positionGroupAbsolute(masuriTabGroup, 5.7281, 18.8218, artboardRect);
            } else if (sponsorPosition == "Middle Sponsor") {
                positionGroupAbsolute(masuriTabGroup, 5.7281, 8.4713, artboardRect);
            }
        }

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
 * Import SVG content into a specific layer of a target document by opening the SVG and copying its active-artboard contents.
 * @param {File} svgFile - The SVG file to open and import.
 * @param {Document} targetDoc - The Illustrator document that will receive the pasted content.
 * @param {Layer} targetLayer - The layer in the target document where the SVG content will be pasted.
 * @throws {Error} If the SVG cannot be opened or the import/paste operation fails; the thrown error contains a descriptive message.
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
 * Group all page items on the given layer into a single group.
 *
 * Clears the document selection before and after grouping; if the layer has no page items nothing is changed.
 *
 * @param {Layer} layer - The layer whose page items will be grouped.
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
 * Aligns the top-left corner of the first group on a layer to the specified coordinates (in centimeters).
 *
 * Positions the group's bounding box so its top-left corner sits at (xCm, yCm). The Y coordinate is measured downward from the document ruler origin (top-left).
 * @param {Layer} layer - Layer whose first groupItem will be positioned; no action if the layer has no groupItems.
 * @param {number} xCm - Target X coordinate in centimeters from the left ruler origin.
 * @param {number} yCm - Target Y coordinate in centimeters from the top ruler origin (measured downward).
 * @throws {Error} If positioning fails.
 */
function positionLayerGroup(layer, xCm, yCm) {
    try {
        // Convert cm to points
        var xPoints = xCm * POINTS_PER_CM;
        var yPoints = yCm * POINTS_PER_CM;

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
 * Resize the given selection uniformly to fit within the specified width and height (centimeters) while preserving aspect ratio.
 *
 * The selection is scaled so its combined bounding box fits inside the target box; the smaller of the width- and height-based scale factors is used. No action is taken for a null or empty selection.
 *
 * @param {Array} selection - Array of page items to scale (e.g., selection array or group items).
 * @param {number} targetWidthCm - Target maximum width in centimeters.
 * @param {number} targetHeightCm - Target maximum height in centimeters.
 * @throws {Error} If resizing fails, an error is thrown with a descriptive message.
 */
function scaleToFit(selection, targetWidthCm, targetHeightCm) {
    try {
        var i;

        if (!selection || selection.length === 0) {
            return;
        }

        // Convert target dimensions from cm to points
        var targetWidthPt = targetWidthCm * POINTS_PER_CM;
        var targetHeightPt = targetHeightCm * POINTS_PER_CM;

        // Get bounding box of selection
        var bounds = selection[0].geometricBounds;
        for (i = 1; i < selection.length; i++) {
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
        for (i = 0; i < selection.length; i++) {
            selection[i].resize(scaleFactor, scaleFactor);
        }

    } catch (e) {
        throw new Error("Failed to scale selection: " + e.message);
    }
}

/**
 * Presents a dialog to choose a Hex color and sponsor position.
 *
 * Shows a configuration window with a color dropdown and a sponsor-position dropdown and returns the selected values, or `null` if the user cancels.
 * @returns {{color: RGBColor, position: string} | null} `color` is an Illustrator `RGBColor` for the selected hex color; `position` is either "Bottom Sponsor" or "Middle Sponsor".
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

    var colorDropdown = colorGroup.add("dropdownlist", undefined, ["Black", "White", "Navy", "Yellow", "Pink", "Red", "Maroon", "Green", "Lime Green", "Emerald Green", "Bottle Green", "Light Blue", "Royal Blue"]);
    colorDropdown.selection = 0; // Default to Black
    colorDropdown.preferredSize.width = 200;

    // Add sponsor position dropdown group
    var positionGroup = dialog.add("group");
    positionGroup.orientation = "row";
    positionGroup.add("statictext", undefined, "Sponsor Position:");

    var positionDropdown = positionGroup.add("dropdownlist", undefined, ["Bottom Sponsor", "Middle Sponsor"]);
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
            rgbColor = hexToRGB("#072441");
        } else if (selectedColor == "Yellow") {
            rgbColor = hexToRGB("#FBBA00");
        } else if (selectedColor == "Pink") {
            rgbColor = hexToRGB("#E14498");
        } else if (selectedColor == "Red") {
            rgbColor = hexToRGB("#C42939");
        } else if (selectedColor == "Maroon") {
            rgbColor = hexToRGB("#621122");
        } else if (selectedColor == "Green") {
            rgbColor = hexToRGB("#6B9B38");
        } else if (selectedColor == "Lime Green") {
            rgbColor = hexToRGB("#70BE46");
        } else if (selectedColor == "Emerald Green") {
            rgbColor = hexToRGB("#00673A");
        } else if (selectedColor == "Bottle Green") {
            rgbColor = hexToRGB("#00523F");
        } else if (selectedColor == "Light Blue") {
            rgbColor = hexToRGB("#00B9ED");
        } else if (selectedColor == "Royal Blue") {
            rgbColor = hexToRGB("#005EA3");
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
 * Sets the fill color for every path on the specified layer.
 *
 * Recursively traverses GroupItem and CompoundPathItem structures, enables fill with the provided RGBColor, and disables stroke on each PathItem.
 * @param {Layer} layer - The Illustrator layer whose path items will be recolored.
 * @param {RGBColor} rgbColor - The RGBColor to apply to each PathItem's fill.
 * @throws {Error} If an error occurs while traversing or modifying layer items.
 */
function applyColorToLayer(layer, rgbColor) {
    try {
        /**
         * Apply the script's `rgbColor` to every PathItem in the given collection, recursing into groups and compound paths.
         *
         * For each PathItem this enables fill, sets its `fillColor` to the outer-scope `rgbColor`, and disables stroke.
         * Traverses GroupItem.pageItems and CompoundPathItem.pathItems recursively.
         *
         * @param {Array} items - An array or collection of page items (PathItem, GroupItem, CompoundPathItem) to process.
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
 * Remove path items in the hex layer whose bounds intersect the sponsor layer's content.
 * This permanently deletes any PathItem or CompoundPathItem from hexLayer that overlaps the combined bounding box of sponsorLayer.
 * @param {Layer} hexLayer - Layer containing hex paths to be checked and removed when overlapping.
 * @param {Layer} sponsorLayer - Layer whose combined content bounds are used to determine overlap.
 */
function removeOverlappingHexPaths(hexLayer, sponsorLayer) {
    try {
        var i;

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
        for (i = 0; i < hexPaths.length; i++) {
            var pathBounds = hexPaths[i].geometricBounds;

            // Check if bounding boxes intersect
            if (boundsIntersect(pathBounds, sponsorBounds)) {
                pathsToDelete.push(hexPaths[i]);
            }
        }

        // Delete overlapping paths
        for (i = 0; i < pathsToDelete.length; i++) {
            pathsToDelete[i].remove();
        }

    } catch (e) {
        throw new Error("Failed to remove overlapping hex paths: " + e.message);
    }
}

/**
 * Compute the combined geometric bounding box for all page items on a layer.
 *
 * @param {Layer} layer - The Illustrator layer whose page items will be inspected.
 * @returns {(number[]|null)} An array `[left, top, right, bottom]` representing the merged geometric bounds of all items on the layer, or `null` if the layer contains no page items.
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
 * Collects all PathItem and CompoundPathItem objects from a list of page items into the provided array.
 * Traverses GroupItem contents recursively and appends each found PathItem or CompoundPathItem to `pathArray`.
 * @param {Array} items - Array or collection of page items to search (PageItem/GroupItem/CompoundPathItem).
 * @param {Array} pathArray - Array to which discovered PathItem and CompoundPathItem objects will be appended.
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
 * @param {number[]} bounds1 - Bounds as [left, top, right, bottom].
 * @param {number[]} bounds2 - Bounds as [left, top, right, bottom].
 * @returns {boolean} `true` if the rectangles defined by the bounds intersect, `false` otherwise.
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
 * Aligns the sponsor layer to the hex layer according to the specified position mode.
 *
 * Positions the sponsor layer relative to the hex layer's bounding box. If either layer has no content, the function does nothing.
 *
 * @param {Layer} sponsorLayer - Layer containing the sponsor artwork to move.
 * @param {Layer} hexLayer - Layer whose bounding box is used as the reference.
 * @param {string} positionType - Position mode: `"Bottom Sponsor"` to horizontally center the sponsor and align its bottom edge with the hex bottom; `"Middle Sponsor"` to center the sponsor both horizontally and vertically within the hex.
 * @throws {Error} If an unexpected error occurs while computing bounds or translating items.
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

        if (positionType == "Bottom Sponsor") {
            // Horizontally centered, bottom edges aligned
            deltaX = hexCenterX - sponsorCenterX;
            deltaY = hexBounds[3] - sponsorBounds[3]; // Align bottom edges
        } else if (positionType == "Middle Sponsor") {
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
 * Remove empty layers from the given document.
 * If a layer cannot be removed, the function skips it without throwing an error.
 * @param {Document} doc - The Illustrator document whose empty layers should be removed.
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
 * Import guide artwork from an SVG into a new "Guides" layer and convert the imported items into Illustrator guides.
 * @param {File} svgFile - The SVG file to open and import.
 * @param {Document} targetDoc - The Illustrator document to receive the imported guides.
 * @returns {Layer} The newly created Guides layer containing the converted guide items.
 * @throws {Error} If the SVG cannot be opened, pasted, translated, or converted to guides.
 */
function importGuidesFromSVG(svgFile, targetDoc) {
    try {
        var i;

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
            var targetX = 0.5 * POINTS_PER_CM;
            var targetY = -0.5 * POINTS_PER_CM;

            // Get bounding box of all pasted items
            if (pastedItems.length > 0) {
                var bounds = pastedItems[0].geometricBounds;
                for (i = 1; i < pastedItems.length; i++) {
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
                for (i = 0; i < pastedItems.length; i++) {
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
 * Recursively convert path items within a collection to Illustrator guides.
 *
 * Traverses the provided items and marks any PathItem encountered as a guide.
 * For GroupItem and CompoundPathItem entries, the function recurses into their child items.
 *
 * @param {Array} items - Array or collection of page items (PathItem, GroupItem, CompoundPathItem) to process.
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
 * Create an Illustrator RGBColor from a hex color string.
 *
 * Accepts 3- or 6-digit hex formats, with or without a leading `#`.
 * @param {String} hex - Hex color code (e.g., "#FFFFFF" or "#FFF").
 * @return {RGBColor} The corresponding Illustrator RGBColor with `red`, `green`, and `blue` components set.
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
 * Group the first group from each provided layer into a single "Artwork" group and center it on the document artboard.
 * Selects the first groupItem from each layer (skipping layers with no groups), groups those items, renames the resulting group to "Artwork", renames its layer to "Artwork" if present, and translates the group so its center aligns with the artboard center.
 * @param {Document} doc - The Illustrator document containing the layers.
 * @param {Array.<Layer>} layers - Array of layers whose first groupItems will be grouped and centered; layers without groupItems are ignored.
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
/**
 * Apply color to an array of page items
 */
function applyColorToItems(items, rgbColor) {
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (item.typename == "PathItem") {
            item.filled = true;
            item.fillColor = rgbColor;
            item.stroked = false;
        } else if (item.typename == "GroupItem") {
            applyColorToItems(item.pageItems, rgbColor);
        } else if (item.typename == "CompoundPathItem") {
            applyColorToItems(item.pathItems, rgbColor);
        }
    }
}

/**
 * Position one group relative to another
 */
function positionGroupRelativeToGroup(sponsorGroup, hexGroup, positionType) {
    var sponsorBounds = sponsorGroup.geometricBounds;
    var hexBounds = hexGroup.geometricBounds;

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

    if (positionType == "Bottom Sponsor") {
        deltaX = hexCenterX - sponsorCenterX;
        deltaY = hexBounds[3] - sponsorBounds[3];
    } else if (positionType == "Middle Sponsor") {
        deltaX = hexCenterX - sponsorCenterX;
        deltaY = hexCenterY - sponsorCenterY;
    }

    sponsorGroup.translate(deltaX, deltaY);
}

/**
 * Remove paths in hexGroup that overlap with sponsorGroup
 */
function removeOverlappingPathsInGroups(hexGroup, sponsorGroup) {
    var sponsorBounds = sponsorGroup.geometricBounds;
    var hexPaths = [];
    collectPathItems(hexGroup.pageItems, hexPaths);

    var pathsToDelete = [];
    for (var i = 0; i < hexPaths.length; i++) {
        var pathBounds = hexPaths[i].geometricBounds;
        if (boundsIntersect(pathBounds, sponsorBounds)) {
            pathsToDelete.push(hexPaths[i]);
        }
    }

    for (var i = 0; i < pathsToDelete.length; i++) {
        pathsToDelete[i].remove();
    }
}

/**
 * Center a group on the artboard
 */
function centerGroupOnArtboard(group, artboardRect) {
    var artboardCenterX = (artboardRect[0] + artboardRect[2]) / 2;
    var artboardCenterY = (artboardRect[1] + artboardRect[3]) / 2;

    var groupBounds = group.geometricBounds;
    var groupCenterX = (groupBounds[0] + groupBounds[2]) / 2;
    var groupCenterY = (groupBounds[1] + groupBounds[3]) / 2;

    var deltaX = artboardCenterX - groupCenterX;
    var deltaY = artboardCenterY - groupCenterY;

    group.translate(deltaX, deltaY);
}

/**
 * Position a group at absolute coordinates relative to artboard's top-left corner
 */
function positionGroupAbsolute(group, xCm, yCm, artboardRect) {
    var xPoints = xCm * POINTS_PER_CM;
    var yPoints = yCm * POINTS_PER_CM;

    // Calculate target position relative to artboard's top-left corner
    var targetX = artboardRect[0] + xPoints;  // artboard left + offset
    var targetY = artboardRect[1] - yPoints;  // artboard top - offset (subtract because Y increases upward)

    var bounds = group.geometricBounds;
    var currentLeft = bounds[0];
    var currentTop = bounds[1];

    var deltaX = targetX - currentLeft;
    var deltaY = targetY - currentTop;

    group.translate(deltaX, deltaY);
}

/**
 * Center two groups together on the artboard without grouping them
 */
function centerTwoGroupsOnArtboard(group1, group2, artboardRect) {
    // Get bounds of both groups
    var bounds1 = group1.geometricBounds;
    var bounds2 = group2.geometricBounds;

    // Calculate combined bounds
    var combinedLeft = Math.min(bounds1[0], bounds2[0]);
    var combinedTop = Math.max(bounds1[1], bounds2[1]);
    var combinedRight = Math.max(bounds1[2], bounds2[2]);
    var combinedBottom = Math.min(bounds1[3], bounds2[3]);

    // Calculate combined center
    var combinedCenterX = (combinedLeft + combinedRight) / 2;
    var combinedCenterY = (combinedTop + combinedBottom) / 2;

    // Calculate artboard center
    var artboardCenterX = (artboardRect[0] + artboardRect[2]) / 2;
    var artboardCenterY = (artboardRect[1] + artboardRect[3]) / 2;

    // Calculate translation needed
    var deltaX = artboardCenterX - combinedCenterX;
    var deltaY = artboardCenterY - combinedCenterY;

    // Translate both groups
    group1.translate(deltaX, deltaY);
    group2.translate(deltaX, deltaY);
}

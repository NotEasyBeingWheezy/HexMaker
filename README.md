# HexMaker

An Adobe Illustrator ExtendScript that automates the creation of hex-patterned sponsor documents for Masuri.

## Overview

HexMaker generates 13cm × 34cm CMYK documents with:
- **Hex pattern layer** with customizable colours
- **Sponsor artwork layer** (copied from active document)
- **Masuri Tab branding layer**
- **Guide layer** for alignment
- Automatic positioning, overlap removal, and grouping

## Features

- **13 Colour Options**: Black, White, Navy, Yellow, Pink, Red, Maroon, Green, Lime Green, Emerald Green, Bottle Green, Light Blue, Royal Blue
- **2 Sponsor Positioning Modes**:
  - Bottom Sponsor (hex at bottom)
  - Middle Sponsor (hex centered)
- **Automatic Overlap Removal**: Removes hex paths that overlap with sponsor content
- **Smart Grouping**: Groups and centers all content into organized "Artwork" group
- **Guide Integration**: Imports and positions guides at 0.5cm from top-left

## Requirements

- Adobe Illustrator CC or later
- Active document with sponsor artwork selected

## Installation

1. Clone or download this repository
2. Ensure the following structure:
   ```
   HexMaker/
   ├── HexMaker.jsx
   └── assets/
       ├── HEX.svg
       ├── MASURI TAB.svg
       └── GUIDES.svg
   ```

## Usage

### Quick Start

1. Open your sponsor artwork in Adobe Illustrator
2. Select all artwork you want to include (guides will be automatically filtered out)
3. Run the script: **File → Scripts → Other Script** → Select `HexMaker.jsx`
4. Choose hex colour and sponsor position from the dialog
5. Click OK

The script will:
- Create a new 13cm × 34cm CMYK document
- Copy your selected artwork to the Sponsor layer (scaled to 11cm × 8cm)
- Import the hex pattern in your chosen colour
- Position everything correctly based on your selected mode
- Remove overlapping hex paths
- Group and center all content

### Configuration Dialog

**Hex Colour**: Choose from 13 preset colours for the hex pattern

**Sponsor Position**:
- **Bottom Sponsor**: Sponsor aligned to bottom of hex pattern
- **Middle Sponsor**: Sponsor centered vertically with hex pattern

## Technical Details

### Document Specifications

- **Size**: 13cm × 34cm (368.504 × 963.78 points)
- **Colour Mode**: CMYK
- **Units**: Centimeters
- **Ruler Origin**: Top-left corner (0, 0)

### Layer Structure

1. **Guides** (locked)
2. **Artwork** (group containing):
   - **Masuri Tab** (positioned at 5.7281cm, 18.8218cm or 8.4713cm)
   - **Hex + Sponsor** (centered group):
     - **Hex** (natural import position with colour applied)
     - **Sponsor** (positioned relative to hex, scaled to fit)

### Positioning Logic

- Hex layer stays at natural SVG import position
- Sponsor positioned relative to hex (bottom-aligned or centered)
- Hex and sponsor grouped and centered on artboard
- Masuri Tab positioned at absolute coordinates
- Final grouping creates "Artwork" master group

## File Descriptions

### HexMaker.jsx
Main ExtendScript file containing all automation logic

### assets/HEX.svg
169 hexagon paths forming the hex pattern template

### assets/MASURI TAB.svg
Masuri branding logo with yellow background

### assets/GUIDES.svg
Guide paths for alignment (converted to Illustrator guides)

## Development

### Code Structure

- **Main workflow**: `main()` function orchestrates the entire process
- **Modular functions**: Import, position, scale, colour, overlap detection
- **Error handling**: Try-catch blocks with user-friendly error messages
- **Constants**: `POINTS_PER_CM = 28.3464567` for unit conversion

### Key Functions

- `showConfigDialog()`: User interface for colour and position selection
- `importSVGByOpening()`: Imports SVG files by opening and copying
- `applyColorToLayer()`: Applies hex colour to all paths recursively
- `positionSponsorLayer()`: Aligns sponsor relative to hex layer
- `removeOverlappingHexPaths()`: Detects and removes overlapping hexagons
- `groupAndCenterLayers()`: Groups and centers content on artboard

## Troubleshooting

**"Please open a document before running this script"**
- Open a document with your sponsor artwork before running the script

**"SVG not found" errors**
- Ensure all SVG files are in the `assets/` folder
- Check file names match exactly: `HEX.svg`, `MASURI TAB.svg`, `GUIDES.svg`

**Sponsor content positioned incorrectly**
- Make sure artwork is selected before running the script
- Verify sponsor artwork fits within 11cm × 8cm bounds

## Contributing

Contributions welcome! Please follow existing code style and test thoroughly with Adobe Illustrator before submitting pull requests.

## License

© 2025 Masuri. All rights reserved.

## Author

Created for Masuri by NotEasyBeingWheezy with assistance from Claude (Anthropic).

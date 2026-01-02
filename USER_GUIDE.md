# HexMaker User Guide

A step-by-step guide to creating hex-patterned sponsor documents for Masuri using the HexMaker script.

## Table of Contents

1. [Before You Start](#before-you-start)
2. [Step-by-Step Instructions](#step-by-step-instructions)
3. [Configuration Options](#configuration-options)
4. [Understanding the Output](#understanding-the-output)
5. [Tips and Best Practices](#tips-and-best-practices)
6. [Common Issues](#common-issues)

---

## Before You Start

### What You Need

- **Adobe Illustrator** (CC or later)
- **HexMaker script** and asset files installed
- **Sponsor artwork** ready in an Illustrator document

### Installation Check

Make sure you have this folder structure:
```
HexMaker/
├── HexMaker.jsx          ← The script file
└── assets/
    ├── HEX.svg           ← Hex pattern template
    ├── MASURI TAB.svg    ← Masuri branding
    └── GUIDES.svg        ← Alignment guides
```

---

## Step-by-Step Instructions

### Step 1: Prepare Your Sponsor Artwork

1. Open Adobe Illustrator
2. Open your sponsor artwork file
3. **Select all the artwork** you want to include in the final document
   - Use **Select → All** or press `Cmd+A` (Mac) / `Ctrl+A` (Windows)
   - Don't worry about guides - they're automatically filtered out!

> **Note**: Your artwork will be scaled to fit within 11cm × 8cm while maintaining proportions

### Step 2: Run the Script

1. Go to **File → Scripts → Other Script...**
2. Navigate to where you saved `HexMaker.jsx`
3. Select `HexMaker.jsx` and click **Open**

### Step 3: Configure Your Document

A dialog box will appear with two options:

#### Hex Colour
Choose the colour for your hex pattern from 13 options:
- Black
- White
- Navy
- Yellow
- Pink
- Red
- Maroon
- Green
- Lime Green
- Emerald Green
- Bottle Green
- Light Blue
- Royal Blue

#### Sponsor Position
Choose how the sponsor content should be positioned:
- **Bottom Sponsor**: Sponsor artwork aligned to bottom of hex pattern
- **Middle Sponsor**: Sponsor artwork centered vertically with hex pattern

### Step 4: Create Document

Click **OK** and the script will:

1. ✓ Create a new 13cm × 34cm CMYK document
2. ✓ Copy your sponsor artwork to the new document
3. ✓ Scale your artwork to fit (max 11cm × 8cm)
4. ✓ Import the hex pattern in your chosen colour
5. ✓ Position your sponsor artwork relative to the hex pattern
6. ✓ Remove any hex shapes that overlap with your sponsor artwork
7. ✓ Import and position the Masuri Tab branding
8. ✓ Import alignment guides
9. ✓ Group everything into an organized "Artwork" group
10. ✓ Center the final artwork on the artboard

When complete, you'll see: **"Hex created successfully!"**

---

## Configuration Options

### Hex Colour Options

All colours are team/brand-specific hex codes:

| Colour       | Hex Code | Use Case |
|--------------|----------|----------|
| Black        | #000000  | Default, high contrast |
| White        | #FFFFFF  | Light backgrounds |
| Navy         | #072441  | Classic team colour |
| Yellow       | #FBBA00  | High visibility |
| Pink         | #E14498  | Specific team branding |
| Red          | #C42939  | Bold, attention-grabbing |
| Maroon       | #621122  | Deep, rich tone |
| Green        | #6B9B38  | Natural, fresh |
| Lime Green   | #70BE46  | Bright, energetic |
| Emerald Green| #00673A  | Premium, classic |
| Bottle Green | #00523F  | Deep, traditional |
| Light Blue   | #00B9ED  | Modern, clean |
| Royal Blue   | #005EA3  | Professional, trustworthy |

### Sponsor Position Modes

#### Bottom Sponsor
- Sponsor artwork aligned to **bottom edge** of hex pattern
- Hex pattern extends above sponsor
- Best for: Sponsors that work well at the base

#### Middle Sponsor
- Sponsor artwork **centered vertically** within hex pattern
- Hex pattern surrounds sponsor on all sides
- Best for: Logos that benefit from centered placement

---

## Understanding the Output

### Document Specifications

Your new document will have:
- **Size**: 13cm wide × 34cm tall
- **Colour Mode**: CMYK (print-ready)
- **Units**: Centimeters
- **Ruler Origin**: Top-left corner at 0, 0

### Layer Structure

Open the **Layers panel** to see:

```
Guides (locked)
└── [Guide objects]

Artwork
├── Masuri Tab
├── Hex
└── Sponsor
```

#### Guides Layer
- Locked to prevent accidental movement
- Contains alignment guides positioned 0.5cm from top-left corner

#### Artwork Layer
Contains a master group with three sub-groups:
- **Masuri Tab**: Branding positioned at specific coordinates
- **Hex**: Coloured hex pattern with overlaps removed
- **Sponsor**: Your scaled sponsor artwork

### What You Can Do Next

- **Export for print**: File → Export → Save for Print
- **Adjust positioning**: Unlock layers and manually adjust if needed
- **Change colours**: Select the Hex group and apply new colours
- **Add additional content**: Work within the Artwork group

---

## Tips and Best Practices

### Getting the Best Results

1. **Prepare Clean Artwork**
   - Remove any unnecessary guides from source document
   - Ensure artwork is properly grouped
   - Check that all elements are visible and unlocked

2. **Choose the Right Mode**
   - Preview your sponsor artwork first
   - Consider which positioning best showcases the sponsor
   - Bottom Sponsor works well for horizontal logos
   - Middle Sponsor works well for square/circular logos

3. **Colour Selection**
   - Choose hex colours that complement (not clash with) sponsor colours
   - Test contrast for readability
   - Consider the final print environment

4. **After Generation**
   - Review overlap removal to ensure nothing important was deleted
   - Check that guides align with your requirements
   - Verify all sponsor details are visible and clear

### Workflow Tips

- **Save your source file first** before running the script
- **Keep the original** - the script creates a new document
- **Run on multiple sponsors** - process several quickly with different settings
- **Create colour variations** - run the same sponsor with different hex colours

---

## Common Issues

### "Please open a document before running this script"

**Problem**: No document is open when you run the script

**Solution**:
1. Open your sponsor artwork file first
2. Make sure the document is active (click on it)
3. Select your artwork
4. Then run the script

---

### "SVG not found" Error

**Problem**: Script can't find the required SVG template files

**Solution**:
1. Check that the `assets/` folder exists in the same location as `HexMaker.jsx`
2. Verify all three files are present:
   - `HEX.svg`
   - `MASURI TAB.svg`
   - `GUIDES.svg`
3. Check file names match exactly (case-sensitive)

---

### Sponsor Content Too Small/Large

**Problem**: Sponsor artwork isn't the right size in final document

**Why**: The script automatically scales sponsor content to fit within 11cm × 8cm maximum

**Solutions**:
- **If too small**: Create larger artwork in your source document before running script
- **If too large**: The script maintains aspect ratio - this is expected behaviour
- **Manual adjustment**: After generation, unlock the Sponsor layer and scale manually

---

### Hex Pattern Looks Wrong

**Problem**: Some hexagons are missing or look incomplete

**Why**: The script removes hex paths that overlap with sponsor content

**Solutions**:
- This is intentional to avoid obscuring sponsor details
- If too many hexagons removed, try different sponsor positioning mode
- Check that sponsor artwork isn't too large relative to the hex pattern

---

### Positioning Doesn't Look Right

**Problem**: Artwork isn't centered or positioned as expected

**Solutions**:
1. Check which positioning mode you selected
2. Unlock the Artwork layer
3. Manually adjust positioning if needed
4. For consistent results, re-run the script with different settings

---

### Can't Edit After Generation

**Problem**: Everything is grouped and hard to edit

**Solution**:
1. Open **Layers panel**
2. Expand the **Artwork** layer
3. Unlock the specific sub-group you want to edit (Hex, Sponsor, or Masuri Tab)
4. Click into the group to edit individual elements
5. Use **Object → Ungroup** if you need to break groups apart

---

## Need Help?

If you encounter issues not covered in this guide:

1. Check the README.md for technical details
2. Verify your Adobe Illustrator version is compatible (CC or later)
3. Ensure all template files are present and unmodified
4. Try running the script with a simple test artwork first

---

## Quick Reference Card

### Running the Script
**File → Scripts → Other Script → Select HexMaker.jsx**

### Requirements Checklist
- [ ] Adobe Illustrator CC or later open
- [ ] Sponsor artwork document active
- [ ] Artwork selected
- [ ] HexMaker.jsx and assets folder in same location

### Configuration Quick Tips
- **Black/White**: Maximum contrast
- **Team colours**: Match sponsor branding
- **Bottom**: Horizontal logos
- **Middle**: Square/circular logos

### Output Specs
- 13cm × 34cm
- CMYK
- Centimeters
- Print-ready

---

*Last updated: January 2025*

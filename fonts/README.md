# Fonts Directory

This directory is for placing **legally obtained** Bengali font files that are required for proper DOCX to PDF conversion.

## Required Fonts

### SutonnyMJ / SutonnyOMJ (Required for Bijoy documents)

If your DOCX files use the **SutonnyMJ** or **SutonnyOMJ** fonts (very common in Bangladeshi government and professional documents), you **must** place the font files here.

**These fonts are copyrighted and NOT included in this repository.**

You need to legally obtain and place the following files:

```
fonts/
├── SutonnyMJ.ttf          (or .otf)
├── SutonnyMJ Bold.ttf     (or .otf)
├── SutonnyOMJ.ttf         (or .otf)
└── SutonnyOMJ Bold.ttf    (or .otf)
```

### Nikosh (Optional)

If your documents use the Nikosh font:

```
fonts/
├── Nikosh.ttf
└── NikoshBAN.ttf
```

## Where to Get These Fonts

- **SutonnyMJ**: Available from [Ananda Computers](http://www.anandacomputers.com/) or included with the Bijoy Bayanno keyboard software.
- **Nikosh**: Available from [Ekushey](https://ekushey.org/) — some versions are open source.

## What Happens Without These Fonts

If a DOCX file uses SutonnyMJ and the font is not installed:

1. The converter will **detect** the legacy font
2. It will **convert the text** from Bijoy encoding to Unicode
3. It will **substitute** with Noto Sans Bengali for rendering
4. The **text will be readable**, but the visual appearance may differ slightly from the original SutonnyMJ rendering

For **pixel-perfect** reproduction of the original document, place the SutonnyMJ font files in this directory.

## Included Fonts (Open Source)

The following Unicode Bengali fonts are installed automatically via the Docker image:

- **Noto Sans Bengali** — Google's open-source Bengali sans-serif font
- **Noto Serif Bengali** — Google's open-source Bengali serif font

These are used as fallback fonts when legacy fonts are not available.

## Installation

When you place font files here and build the Docker image, they are automatically:

1. Copied to `/usr/local/share/fonts/`
2. Registered with `fc-cache -fv`
3. Available to LibreOffice for PDF rendering

For local development, install the fonts system-wide:

- **Linux**: Copy to `~/.fonts/` or `/usr/local/share/fonts/` and run `fc-cache -fv`
- **macOS**: Double-click each font file to install via Font Book
- **Windows**: Right-click each font file → Install

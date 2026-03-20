# OOXML Fallback

Use this fallback only when native PowerPoint COM cannot express the required change safely.

Good fallback cases:

- repair or inspect broken package internals
- remove orphaned OOXML resources
- duplicate or create slides by editing package parts
- validate a repaired deck after XML-level edits

## Bundled utilities

### Unpack a presentation
```bash
python scripts/office/unpack.py input.pptx unpacked
```

### Add a slide or duplicate one
```bash
python scripts/add_slide.py unpacked slide2.xml
```

### Clean orphaned resources
```bash
python scripts/clean.py unpacked
```

### Validate package structure
```bash
python scripts/office/validate.py unpacked --original input.pptx
```

### Pack a new `.pptx`
```bash
python scripts/office/pack.py unpacked repaired.pptx --original input.pptx
```

## Rules

- Prefer COM first. Do not unpack unless there is a clear reason.
- After OOXML edits, always repack, reopen in PowerPoint, and export PNGs for QA.
- Keep this path as a surgical fallback, not the default workflow.

# Spreadsheet standards

Apply these standards unless the user instructs otherwise or an existing workbook already has its own conventions.

## All Excel files

### Professional font

Use a consistent professional font such as Arial or Times New Roman unless the workbook already uses another established font.

### Zero formula errors

Every workbook should be delivered with zero visible Excel error values, including:

- `#REF!`
- `#DIV/0!`
- `#VALUE!`
- `#N/A`
- `#NAME?`
- `#NULL!`
- `#NUM!`
- `#SPILL!`
- `#CALC!`

### Preserve existing templates

When updating an existing template:

- study the workbook before changing it
- match the workbook's existing style and layout
- do not impose a generic house style on a workbook with established conventions

## Financial models

### Color coding standards

Unless the workbook already uses a different system or the user specifies otherwise:

- Blue text `RGB(0,0,255)`: hardcoded inputs and scenario inputs
- Black text `RGB(0,0,0)`: formulas and calculations
- Green text `RGB(0,128,0)`: links to other worksheets in the same workbook
- Red text `RGB(255,0,0)`: external links to other files
- Yellow fill `RGB(255,255,0)`: key assumptions that need user attention or update

### Number formatting standards

- Years: format as text-like year labels such as `2024`, not `2,024`
- Currency: use formats like `$#,##0` and specify units in headers when helpful, such as `Revenue ($mm)`
- Zeros: display as `-` using custom number formats when appropriate
- Percentages: default to `0.0%`
- Multiples: use formats like `0.0x`
- Negative numbers: use parentheses instead of a minus sign when that matches finance conventions

### Formula construction rules

#### Assumptions placement

- place assumptions such as growth rates, margins, and multiples in dedicated input cells
- reference those cells from formulas
- avoid embedding hardcoded constants directly in formulas unless the user explicitly wants that

Example:

- Prefer `=B5*(1+$B$6)`
- Avoid `=B5*1.05`

#### Error prevention

- verify every reference after row or column insertions
- check range bounds to avoid off-by-one mistakes
- keep formulas consistent across repeated periods
- test zero, negative, and large-value cases when relevant
- avoid unintended circular references

### Documentation for hardcodes

For important hardcoded assumptions or manually entered sourced figures, add cell comments or adjacent notes when the workbook style supports it.

Suggested format:

`Source: [System or Document], [Date], [Specific Reference], [URL if applicable]`

Examples:

- `Source: Company 10-K, FY2024, Page 45, Revenue Note, [SEC EDGAR URL]`
- `Source: Company 10-Q, Q2 2025, Exhibit 99.1, [SEC EDGAR URL]`
- `Source: Bloomberg Terminal, 8/15/2025, AAPL US Equity`
- `Source: FactSet, 8/20/2025, Consensus Estimates Screen`

## Formula verification checklist

### Essential checks

- test two or three sample references before filling across the model
- confirm the intended Excel columns and rows are being used
- remember Excel is one-indexed

### Common pitfalls

- null or NaN values causing `#VALUE!`
- division by zero causing `#DIV/0!`
- bad references causing `#REF!`
- wrong sheet names in cross-sheet formulas
- formulas copied across sections without adjusting anchors correctly
- far-right columns being misidentified during programmatic generation

### Testing strategy

- start with a small sample before broad propagation
- verify all referenced cells exist
- test representative edge cases

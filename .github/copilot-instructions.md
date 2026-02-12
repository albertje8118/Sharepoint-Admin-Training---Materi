# Copilot Instructions — SharePoint Admin Training Course

## Project Overview
A 3-day instructor-led training course (12 modules) for SharePoint Online administrators, following Microsoft Official Curriculum (MOC) style. Lab scenario: "Project Northwind Intranet Modernization" with 10 participants (P01–P10) + 1 trainer on a shared M365 tenant.

## Architecture & Content Pipeline
- **Markdown is source of truth** — `Module-XX-Slides.md` defines slide content/speaker notes; `Module-XX-Student-Guide.md` and `Lab-XX-*.md` are the reading/lab materials.
- **Python generators** (`generate_moduleNN_slides.py`) read nothing — all content is hardcoded inside each script using `python-pptx`. They output `Module-XX-Slides.pptx`.
- **Git tracks only `.md` and `.pptx`** — all `.py` and `.ps1` files are in `.gitignore`. Generators live in the repo working tree but are not committed.
- `participant-packs/generate_participant_packs.py` produces per-participant DOCX/PDF/PPTX/XLSX packs using `python-pptx`, `python-docx`, `openpyxl`, and `reportlab`.

## Slide Generator Conventions (Original Design)
Every module generator is a **standalone script** — no shared library. Each file copy-pastes the same helpers and palette. When creating or editing a generator:

- **Slide size**: `Inches(13.333) × Inches(7.5)` (16:9 widescreen)
- **Layout**: Always use `prs.slide_layouts[6]` (blank layout)
- **Color palette**: `DARK_BG=#1B1B2F`, `ACCENT_BLUE=#0078D4`, `ACCENT_TEAL=#00B294`, `ACCENT_PURPLE=#6B69D6`, `ORANGE=#FF8C00`, `GREEN=#107C10`
- **Font**: `Segoe UI` for all text; bullet icon is `▸`
- **Required helpers**: `add_solid_bg`, `add_shape_rect`, `add_rounded_rect`, `add_text_box`, `add_bullet_frame`, `add_speaker_notes`, `add_top_bar`, `add_footer_bar`, `section_divider`, `new_slide`
- **Footer pattern**: dark bar with `"Module N | Title"` left + `"slide_num / total"` right
- **Output path**: `os.path.join(os.path.dirname(__file__), "Module-NN-Slides.pptx")`
- **Track slide count** via `slide_counter = [0]` (mutable list) and `TOTAL_SLIDES`

## Alternative Design (`module 1-alt/`)
A second visual style called "Clean Minimal with Bold Sidebar" lives in `module 1-alt/`. Key differences:
- Warm off-white background (`#FAF8F5`), earth-tone accent palette (coral `#E85D4A`, indigo `#2D316F`, warm teal `#009B8D`, golden `#E5A100`)
- Bold left-sidebar accent strip instead of top bar; pill-shaped badges; progress dots in footer
- Font: `Calibri` for headings, `Segoe UI Semilight` for body
- Additional helpers: `add_pill`, `add_left_sidebar`, `add_progress_footer`, `add_multiline`

When asked to create a new design variant, place it in a `module N-alt/` folder with its own generator.

## File Naming
```
module N/
  generate_moduleNN_slides.py    # NN = zero-padded module number
  Module-NN-Slides.md            # slide content outline
  Module-NN-Slides.pptx          # generated output
  Module-NN-Student-Guide.md     # student reading
  Lab-NN-<Descriptive-Title>.md  # lab instructions
  README.md                      # module overview
```

## Lab & Scenario Rules
- Participant naming: `NW-Pxx-<Purpose>` (e.g., `NW-P03-ProjectSite`)
- Labs should be **verification-first** — avoid tenant-wide changes unless marked Trainer-only
- Always specify estimated time, lab type (UI/PowerShell), deliverables, and required roles
- Reference `../scenario/Lab-Scenario-Overview.md` for shared-tenant context

## Running Generators
```powershell
# From workspace root (Conda env with python-pptx installed):
C:/Users/alber/anaconda3/Scripts/conda.exe run -p C:\Users\alber\anaconda3 --no-capture-output python "module 1\generate_module01_slides.py"
```

## PowerShell Scripts (`scripts/`)
- `Provision-TrainingUsers-P01-P10.ps1` — creates P01–P10 accounts via Microsoft Graph PowerShell
- `Reassign-Licenses-TrainingUsers.ps1` — two-phase license reset + assignment; supports `-WhatIf`
- Both require `Microsoft.Graph` module and use `Set-StrictMode -Version Latest`

## Key Constraints
- Never hardcode service limit numbers in slides — teach "verify in Microsoft Learn"
- Teach "how to find settings" rather than "click exactly here" (UI varies by tenant)
- Each generator is self-contained; do not extract shared modules (this is intentional for portability)
- Content must align with 2026 Microsoft 365 service state

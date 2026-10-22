# resume

Zachary Yaninek's resume on the MIT Sloan MBAn 2027 template.

- `resume.json`: resume content (no phone number)
- `template/MBAn-Resume_Template_2027.docx`: the MIT Sloan CDO template; its styles, bullets, fonts and margins are reused as is
- `build_resume.py`: regenerates the template's document body from `resume.json`
- `Yaninek_Zachary_MBAn2027_Resume.docx` / `.pdf`: public build, phone omitted

## Build

```bash
# public build (no phone), written next to this README
python3 build_resume.py --pdf

# full build for submission; keep it out of this public repo
python3 build_resume.py --phone XXX-XXX-XXXX --out ~/private/Yaninek_Zachary_MBAn2027_Resume.docx --pdf
```

PDF conversion uses LibreOffice. Without Times New Roman installed it substitutes the metric-compatible Liberation Serif, so line breaks match Word. Export the PDF you submit to Career Central from Word.

Layout knobs live under `layout` in `resume.json`: `section_gap` and `entry_gap`, both in twips (1/20 pt).

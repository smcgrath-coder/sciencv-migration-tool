# SciENcv Biosketch Migration Tool

**Addressing the pain of transferring existing biosketches into SciENcv**

Starting January 25, 2026, NIH requires all biosketches to be submitted via SciENcv using the new Common Forms. This tool helps researchers convert their existing Word/PDF biosketches into a format ready for SciENcv entry.

## The Problem

As [discussed on social media](https://bsky.app/profile/abepalmer.bsky.social/post/3mbs47ohmgc2n):

> "If you are submitting an NIH grant in February, you will be required to use SciENcv to prepare you biosketch. IT IS MUCH WORSE THAN YOU CAN POSSIBLY IMAGINE. Set aside *at least* 4 hours just to transfer an existing biosketch into SciENcv."

**Key pain points:**
1. **No Word/PDF upload for biosketches** - While SciENcv can export XML, it **cannot import biosketch XML**. Only Current & Pending (Other) Support documents accept XML upload.
2. **Inconsistent Markdown handling** - Personal Statement uses Enter for line breaks, but Contributions require TWO trailing spaces
3. **Character limits** - New Common Forms have strict limits (Personal Statement: 3500, Contributions: 2000 each)
4. **Manual data entry** - Everything must be copied/pasted or re-entered for your FIRST biosketch

**The Good News (as of December 2025):**
Once you have ONE biosketch in SciENcv, creating subsequent ones is much easier. The new Common Forms allow Personal Statement, Contributions to Science, Honors, and Products to all import from existing SciENcv biosketches. **This tool helps you get over that initial hurdle.**

## Important: Understanding SciENcv XML Capabilities

| Document Type | Can Export XML? | Can Import XML? |
|--------------|-----------------|-----------------|
| NIH Biosketch | ✅ Yes | ❌ **NO** |
| NSF Biosketch | ✅ Yes | ❌ **NO** |
| Current & Pending (Other) Support | ✅ Yes | ✅ **YES** |

This is why the transition is so painful - you CAN'T just convert your Word biosketch to XML and upload it. You must manually copy/paste each section. This tool helps by:
1. Parsing your existing biosketch into SciENcv-ready sections
2. Adding proper Markdown formatting (especially the two-space line breaks)
3. Validating character counts before you start

## What This Tool Does

### ✅ Biosketch Conversion (Word/PDF → Copy-Paste Ready)
- Parses your existing biosketch and extracts each section
- Converts text to SciENcv-compatible Markdown format
- **Automatically adds two trailing spaces** for line breaks in Contributions
- Validates character counts against Common Form limits
- Generates copy-paste ready text blocks for each SciENcv field

### ✅ Current & Pending XML Generation (Direct Upload!)
- Converts Other Support entries to valid SciENcv XML
- This XML **CAN be directly uploaded** to SciENcv
- Saves significant time on the most tedious part

## Installation

### Option 1: Use with Claude Code
```bash
# Clone or download the script
curl -O https://raw.githubusercontent.com/[your-repo]/sciencv_converter.py

# Run Claude Code and ask it to help you convert your biosketch
claude "Please convert my biosketch to SciENcv format" --file my_biosketch.docx
```

### Option 2: Standalone Python Script
```bash
# Install dependencies
pip install python-docx PyPDF2

# Run the converter
python sciencv_converter.py my_biosketch.docx
python sciencv_converter.py my_biosketch.pdf
```

## Usage

### Converting a Biosketch

```python
from sciencv_converter import convert_biosketch

# Convert Word document
results = convert_biosketch("my_biosketch.docx")

# Convert PDF
results = convert_biosketch("my_biosketch.pdf", output_dir="./converted")
```

**Output:** A text file with copy-paste ready sections:

```
============================================================
SECTION: Contribution 1
============================================================
Character count: 761/2000 ✅

--- COPY BELOW THIS LINE ---

My early work focused on establishing the role of microRNAs...  
[Note: trailing spaces added automatically for line breaks]

--- END COPY ---
```

### Generating Current & Pending XML (uploadable!)

```python
from sciencv_converter import convert_other_support_to_xml

entries = [
    {
        'contributiontype': 'award',
        'title': 'MicroRNA therapeutics for heart failure',
        'status': 'active',
        'sponsor': 'NIH/NHLBI',
        'award_number': 'R01 HL123456',
        'startdate': '2023-09-01',
        'enddate': '2028-08-31',
        'personmonths': 3.6,
        'overallobjectives': 'The goal of this project is to develop...',
        'pi_name': 'John Smith, PhD'
    },
    # Add more entries...
]

convert_other_support_to_xml(entries, "my_support.xml")
# Upload this directly to SciENcv!
```

## SciENcv Markdown Reference

SciENcv uses Markdown, but with **inconsistent behavior**:

| Section | Line Breaks | Bold | Italic |
|---------|-------------|------|--------|
| Personal Statement | Enter works | `**text**` | `*text*` |
| Contributions to Science | Two spaces + Enter | `**text**` | `*text*` |

**⚠️ HTML tags are NOT supported**

### Line Break Example (Contributions)

Wrong (won't work):
```
First paragraph.
Second paragraph.
```

Correct (note the two spaces at end of line 1):
```
First paragraph.  
Second paragraph.
```

## Character Limits (NIH Common Forms - Required Jan 25, 2026)

| Section | Limit | Notes |
|---------|-------|-------|
| Personal Statement | 3,500 characters | Single statement for all biosketches |
| Each Contribution to Science | 2,000 characters | Max 5 contributions |
| Overall Objectives (Other Support) | 1,500 characters | Per project/proposal |
| Honors | Max 15 entries | No character limit per entry |

**Note:** The Common Forms are **required** for all NIH applications with due dates on or after January 25, 2026. Word-generated biosketches are no longer accepted.

## Workflow Recommendations

### Strategy 1: ORCID-First (Recommended for Future Use)
If you have time before your deadline, populating ORCID is the best long-term strategy:

1. **Set up/update your ORCID** at orcid.org
2. **Add your publications** - ORCID can pull from PubMed, CrossRef, etc.
3. **Add employment history, education, awards**
4. **Link ORCID to My NCBI** in Account Settings
5. **Create SciENcv biosketch** using "External Source: ORCID"

SciENcv will auto-populate much of your data. The ORCID data persists, making future biosketches easier.

### Strategy 2: This Tool (For Immediate Migration)
If you have a deadline and an existing Word/PDF biosketch:

1. Run this converter on your existing biosketch
2. Copy/paste each section into SciENcv
3. Add citations from My Bibliography
4. Your first biosketch becomes a template for all future ones

### Before the Deadline Rush

1. **Set up your ORCID** - Link it to My NCBI/eRA Commons
2. **Populate My Bibliography** - Your publications should be there
3. **Run this converter** on your existing biosketch
4. **Review character counts** - Trim sections that exceed limits

### Day of Conversion

1. Log into SciENcv via My NCBI or eRA Commons
2. Create new "NIH Biographical Sketch Common Form"
3. Use the converter output to copy/paste each section
4. Add citations from My Bibliography
5. For Other Support: Upload the generated XML directly!

## Known Limitations

- PDF parsing quality depends on the PDF structure
- Complex table formatting may not parse perfectly
- Always review converted content before submission
- Education/Training still needs manual entry in SciENcv forms

## Contributing

Found a bug? Have an improvement? 

This tool was created to help the research community navigate the SciENcv transition. Contributions welcome!

## References

- [SciENcv Official Site](https://www.ncbi.nlm.nih.gov/sciencv/)
- [NIH Common Forms Notice](https://grants.nih.gov/grants/guide/notice-files/NOT-OD-26-018.html)
- [SciENcv XML Documentation](https://support.nlm.nih.gov/kbArticle/?pn=KA-05499)
- [UCI's ZotGPT for Other Support](https://research.bio.uci.edu/sciencv/) - Similar approach for C&P forms

## License

MIT - Use freely, help your colleagues!

---

*Created in response to the collective frustration of researchers everywhere facing the January 2026 SciENcv deadline.*

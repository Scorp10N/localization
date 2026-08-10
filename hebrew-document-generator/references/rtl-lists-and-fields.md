# RTL-correct numbered lists and cross-reference fields (DOCX)

Referenced from `SKILL.md` Step 5. This code is adapted from a working,
audit-verified generator (structural checks: every Hebrew paragraph has
`<w:bidi/>`, no run in a mixed paragraph is `<w:rtl/>`-flagged, every run has
`w:cs`/`w:szCs`, `numbering.xml` has all `abstractNum` before all `num`, and
`pPr` children follow the CT_PPr schema order) — not hand-derived from the
spec without execution. Two OOXML schema-order rules matter throughout and
are easy to violate silently (LibreOffice/Chromium both tolerate the
violation and render it correctly anyway, which is exactly why this needs
stating explicitly rather than trusting a visual check):

- **`CT_Numbering` requires every `<w:abstractNum>` before every `<w:num>`.**
  Appending a new `abstractNum` after the document's existing `<w:num>`
  elements (which is what a plain `.append()` does) produces an invalid
  `numbering.xml`.
- **`CT_PPr` requires `<w:numPr>` (schema position 6) before `<w:bidi/>`**
  (position 18), which itself comes before `<w:ind>` (22) and `<w:jc>` (26).
  Appending `bidi` and then `numPr` — the natural order if you write the
  "set RTL" line first — produces an out-of-order `pPr`.

## Shared setup

```python
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def _alloc_numbering_ids(numbering_elem):
    """Next free abstractNumId/numId. Do not hardcode an id and hope it's
    free — the stock template already uses abstractNumId 0-8 / numId 1-9,
    and a corporate .dotx or a second custom definition can occupy anything.
    """
    existing_abstract = [int(el.get(qn('w:abstractNumId')))
                          for el in numbering_elem.findall(qn('w:abstractNum'))]
    existing_num = [int(el.get(qn('w:numId')))
                    for el in numbering_elem.findall(qn('w:num'))]
    return max(existing_abstract, default=-1) + 1, max(existing_num, default=0) + 1

def _register_numbering_def(numbering_elem, abstractNum, num):
    """Insert respecting CT_Numbering's required child order (see above):
    numPicBullet*, abstractNum*, num*, numIdMacAtCleanup?. Anchor the new
    abstractNum before the first existing <w:num>, and the new num before
    <w:numIdMacAtCleanup> if the document happens to have one — a plain
    .append() for either would silently land after it."""
    first_num = numbering_elem.find(qn('w:num'))
    if first_num is not None:
        first_num.addprevious(abstractNum)
    else:
        numbering_elem.append(abstractNum)
    cleanup = numbering_elem.find(qn('w:numIdMacAtCleanup'))
    if cleanup is not None:
        cleanup.addprevious(num)
    else:
        numbering_elem.append(num)

_PPR_TAG_SEQ = [
    'w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore', 'w:framePr',
    'w:widowControl', 'w:numPr', 'w:suppressLineNumbers', 'w:pBdr', 'w:shd',
    'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap',
    'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN',
    'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind',
    'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc',
    'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap', 'w:outlineLvl',
    'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange',
]

def _insert_in_pPr_order(pPr, new_el, tag):
    """Insert new_el at the position CT_PPr's schema sequence requires,
    instead of a plain .append(). python-docx's own accessors (pStyle,
    numPr, tabs, spacing, ind, jc, outlineLvl, sectPr) already insert in
    order; this covers elements built by hand — bidi in particular.
    Idempotent: most CT_PPr children (bidi included) are maxOccurs=1, so if
    `tag` is already present this is a no-op rather than a silent duplicate
    (which real Word will reject/repair) — do not call this from a path
    that also independently sets the same property another way."""
    my_idx = _PPR_TAG_SEQ.index(tag)
    for child in pPr:
        child_tag = 'w:' + child.tag.split('}')[-1]
        if child_tag == tag:
            return
        if child_tag in _PPR_TAG_SEQ and _PPR_TAG_SEQ.index(child_tag) > my_idx:
            child.addprevious(new_el)
            return
    pPr.append(new_el)
```

## RTL bullet numbering

```python
def add_rtl_bullet_numbering(doc):
    """Register a custom bullet numbering definition with RTL-mirrored
    indentation, and return its numId. Call once per document. Do NOT
    reuse python-docx's built-in 'List Bullet' style — its numbering
    definition hardcodes <w:lvlJc w:val="left"/> and
    <w:ind w:left="1440" w:hanging="360"/> (a physical LEFT hanging indent
    with no RTL counterpart) inside a level whose own pPr looks like:
      <w:lvl><w:start .../><w:numFmt .../><w:lvlText .../><w:lvlJc .../>
        <w:pPr><w:tabs>...</w:tabs><w:ind w:left="1440" w:hanging="360"/></w:pPr>
      </w:lvl>
    — note w:ind lives inside w:lvl/w:pPr, not as a direct child of w:lvl.
    """
    # doc.part.numbering_part raises NotImplementedError if the document
    # has never had a numbering part (python-docx's fallback is an
    # unimplemented NumberingPart.new()) — a template that's never used a
    # list is exactly the "corporate .dotx" case above. Force the part into
    # existence first if that's a risk: add and remove one throwaway
    # paragraph with style 'List Bullet', then proceed.
    numbering_part = doc.part.numbering_part
    numbering_elem = numbering_part.element
    abstract_id, num_id = _alloc_numbering_ids(numbering_elem)

    abstractNum = OxmlElement('w:abstractNum')
    abstractNum.set(qn('w:abstractNumId'), str(abstract_id))
    lvl = OxmlElement('w:lvl')
    lvl.set(qn('w:ilvl'), '0')

    start = OxmlElement('w:start'); start.set(qn('w:val'), '1')
    numFmt = OxmlElement('w:numFmt'); numFmt.set(qn('w:val'), 'bullet')
    lvlText = OxmlElement('w:lvlText'); lvlText.set(qn('w:val'), '•')
    lvlJc = OxmlElement('w:lvlJc'); lvlJc.set(qn('w:val'), 'right')

    lvlPPr = OxmlElement('w:pPr')
    lvlPPr.append(OxmlElement('w:bidi'))  # level default is RTL base
    ind = OxmlElement('w:ind')
    ind.set(qn('w:right'), '432')    # chosen right-side hang, not a literal
    ind.set(qn('w:hanging'), '288')  # mirror of the built-in's 1440/360
    lvlPPr.append(ind)

    lvlRPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'David'); rFonts.set(qn('w:hAnsi'), 'David'); rFonts.set(qn('w:cs'), 'David')
    lvlRPr.append(rFonts)

    for el in (start, numFmt, lvlText, lvlJc, lvlPPr, lvlRPr):
        lvl.append(el)
    abstractNum.append(lvl)

    num = OxmlElement('w:num')
    num.set(qn('w:numId'), str(num_id))
    abstractNumId_ref = OxmlElement('w:abstractNumId')
    abstractNumId_ref.set(qn('w:val'), str(abstract_id))
    num.append(abstractNumId_ref)

    _register_numbering_def(numbering_elem, abstractNum, num)
    return num_id

def add_rtl_bullet(doc, num_id, text, font='David', size=11, bold_prefix=None):
    """Real numPr-backed bullet (not a typed '•' run) with RTL-mirrored
    geometry. This helper always sets bidi, on the assumption that bullets
    in a Hebrew document are Hebrew (or mixed) content; for a bullet list
    that might be pure English, check `_para_is_rtl(text)` first the same
    way `add_rtl_paragraph` does, per rule 1 in Step 5."""
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    numPr = OxmlElement('w:numPr')
    ilvl = OxmlElement('w:ilvl'); ilvl.set(qn('w:val'), '0')
    numIdEl = OxmlElement('w:numId'); numIdEl.set(qn('w:val'), str(num_id))
    numPr.append(ilvl); numPr.append(numIdEl)
    _insert_in_pPr_order(pPr, numPr, 'w:numPr')          # numPr before bidi —
    _insert_in_pPr_order(pPr, OxmlElement('w:bidi'), 'w:bidi')  # required order
    # Alignment deliberately left UNSET — same reasoning as table cells in
    # "Hebrew Tables in DOCX": the paragraph's visual position is governed
    # by the numbering definition's own indent, a separate positioning
    # subsystem from body-text w:jc. An explicit physical RIGHT is the one
    # place proven (via the table-cell bug) to fight a subsystem-driven
    # mirror; leaving it unset lets <w:bidi/> + the numbering's own
    # right-indent place it correctly.
    if bold_prefix:
        _add_runs(p, bold_prefix, font=font, size=size, bold=True)
        _add_runs(p, text, font=font, size=size)
    else:
        _add_runs(p, text, font=font, size=size)
    return p
```

The same hardcoded-LTR problem exists for **heading auto-numbering**: linking a multilevel numbering definition to `Heading1`/`Heading2` via `<w:pStyle>` (Word's native "numbered headings" feature, instead of typing `"1. "`/`"2.1 "` literally into the heading text) needs the identical RTL treatment. Add one `<w:lvl>` per heading depth to the `abstractNum` above, each with a `<w:pStyle w:val="Heading1"/>` (or `Heading2`, etc. — `w:pStyle` goes after `w:numFmt`, before `w:lvlText`) and a `%1.` / `%1.%2.` `lvlText` pattern, and use `_register_numbering_def`/`_insert_in_pPr_order` exactly as above. Number the paragraph via `<w:numPr>` on each heading (don't rely on `pStyle`-linking alone to reach every renderer) and remove the typed number from the heading text — Word supplies it.

## Bookmarks and REF-field citations

Once headings use real auto-numbering, any inline citation of a section number in body text — "see section 5.2", or the Hebrew equivalent `(§5.2)` — needs to be a **field**, not typed text, or it silently goes stale the moment a section is added, removed, or reordered.

```python
_BOOKMARK_ID_COUNTER = 0  # bookmark ids must be unique document-wide; don't
                          # take one as a caller-supplied param and hope
                          # nothing else in the document already used it —
                          # same "don't hardcode an id" principle as
                          # _alloc_numbering_ids above.

def bookmark_paragraph(paragraph, name):
    """bookmarkStart/End are CT_P block-content siblings that come AFTER
    pPr (pPr, when present, must be the paragraph's first child) —
    inserting at index 0 unconditionally would place bookmarkStart BEFORE
    pPr and produce an invalid document."""
    global _BOOKMARK_ID_COUNTER
    bookmark_id = _BOOKMARK_ID_COUNTER
    _BOOKMARK_ID_COUNTER += 1
    start = OxmlElement('w:bookmarkStart')
    start.set(qn('w:id'), str(bookmark_id))
    start.set(qn('w:name'), name)
    end = OxmlElement('w:bookmarkEnd')
    end.set(qn('w:id'), str(bookmark_id))
    pPr = paragraph._p.find(qn('w:pPr'))
    insert_at = list(paragraph._p).index(pPr) + 1 if pPr is not None else 0
    paragraph._p.insert(insert_at, start)
    paragraph._p.append(end)

def _field_rpr(run, font='David', size=11):
    """Field-carrier runs need w:cs/w:szCs like any other run (rule 4) —
    easy to forget since they often have no visible w:t."""
    rPr = run._r.get_or_add_rPr()
    rPr.append(rPr.makeelement(qn('w:rFonts'), {
        qn('w:ascii'): font, qn('w:hAnsi'): font, qn('w:cs'): font}))
    rPr.append(rPr.makeelement(qn('w:sz'), {qn('w:val'): str(size * 2)}))
    rPr.append(rPr.makeelement(qn('w:szCs'), {qn('w:val'): str(size * 2)}))

def add_ref_field(paragraph, bookmark_name, font='David', size=11):
    """REF field (\\r = number only, \\h = hyperlinked). Five explicit
    statements rather than a tag-dispatch loop — a loop form is no harder
    to get schema-correct, but is harder for a reader to verify by eye,
    and eyeballing is exactly what catches the two bugs above."""
    r_begin = paragraph.add_run()
    fld_begin = OxmlElement('w:fldChar'); fld_begin.set(qn('w:fldCharType'), 'begin')
    r_begin._r.append(fld_begin); _field_rpr(r_begin, font, size)

    r_instr = paragraph.add_run()
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve')
    instr.text = f' REF {bookmark_name} \\r \\h '
    r_instr._r.append(instr); _field_rpr(r_instr, font, size)

    r_sep = paragraph.add_run()
    fld_sep = OxmlElement('w:fldChar'); fld_sep.set(qn('w:fldCharType'), 'separate')
    r_sep._r.append(fld_sep); _field_rpr(r_sep, font, size)

    r_result = paragraph.add_run('…')  # last-calculated-value placeholder
    _field_rpr(r_result, font, size)

    r_end = paragraph.add_run()
    fld_end = OxmlElement('w:fldChar'); fld_end.set(qn('w:fldCharType'), 'end')
    r_end._r.append(fld_end); _field_rpr(r_end, font, size)
```

## Mid-sentence citations

If a citation sits mid-sentence in a paragraph you'd otherwise build with `add_rtl_paragraph`, split the string at the citation and build the paragraph from parts — but **compute "does this paragraph contain Latin?" once from the full concatenated text, not per fragment.** A trailing fragment like `').'` has no Latin letters *in itself*; judged in isolation it looks like a pure-Hebrew/neutral run and gets wrongly `<w:rtl/>`-flagged even though a sibling fragment in the same paragraph (e.g. `'WinRM'`) makes the paragraph mixed overall — which per rule 2 means no run in it should carry `<w:rtl/>`. `_add_runs`'s `para_has_latin` parameter (Shared setup, Step 5) exists exactly for this:

```python
class Ref:
    """Sentinel marking a REF-field citation inside a parts list — e.g.
    ['...ראשונית (§', Ref('sec5_2'), ' — gMSA...']. Only the number becomes
    a field; surrounding punctuation stays literal text."""
    def __init__(self, bookmark):
        self.bookmark = bookmark

def add_rtl_paragraph_parts(doc, parts, font='David', size=12, heading_level=None):
    """Like add_rtl_paragraph, but parts is a list of str / Ref, for a
    paragraph that embeds one or more REF-field citations mid-sentence."""
    full_text = ''.join(p for p in parts if isinstance(p, str))
    para_has_latin = any(ch.isascii() and ch.isalpha() for ch in full_text)
    p = doc.add_heading(level=heading_level) if heading_level else doc.add_paragraph()
    base_rtl = _para_is_rtl(full_text)
    pPr = p._p.get_or_add_pPr()
    if base_rtl:
        pPr.append(pPr.makeelement(qn('w:bidi'), {}))
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if base_rtl else WD_ALIGN_PARAGRAPH.LEFT
    for part in parts:
        if isinstance(part, Ref):
            add_ref_field(p, part.bookmark, font=font, size=size)
        else:
            _add_runs(p, part, font=font, size=size, para_has_latin=para_has_latin)
    return p
```

**Known limitation, not a bug to chase:** `REF`/`TOC` fields display their *last-calculated value* until explicitly updated. Real Word updates them automatically on open in normal interactive use (or via F9 / right-click → Update Field). Headless conversion (`soffice --headless --convert-to pdf`) was expected to skip that update pass — in practice, testing against LibreOffice found it *does* resolve `REF` fields on conversion (unlike `TOC` fields, which stay as whatever placeholder text you seed them with). Don't rely on that either way across LibreOffice versions; verify field resolution by opening interactively when it matters.

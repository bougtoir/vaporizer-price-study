"""
Generate JATM (Journal of Anesthesia and Translational Medicine) cover letter
as editable .docx file.
"""
import os
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
out_dir = os.path.join(SCRIPT_DIR, 'papers')
os.makedirs(out_dir, exist_ok=True)


def write_cover_letter():
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    style.paragraph_format.line_spacing = 1.15
    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(2.54)
        section.right_margin = Cm(2.54)

    # Date
    from datetime import date
    doc.add_paragraph(date.today().strftime('%B %d, %Y'))
    doc.add_paragraph()

    # Addressee
    doc.add_paragraph('Editor-in-Chief')
    doc.add_paragraph('Journal of Anesthesia and Translational Medicine')
    doc.add_paragraph()

    # Salutation
    doc.add_paragraph('Dear Editor,')
    doc.add_paragraph()

    # Body
    doc.add_paragraph(
        'We are resubmitting our revised manuscript entitled "Targeted environmental regulation '
        'and secondary market vaporizer prices: an observational analysis of the EU desflurane '
        'phase-out" (Manuscript Number JATMED-D-26-00064) for reconsideration as an Article in '
        'the Journal of Anesthesia and Translational Medicine.')

    doc.add_paragraph(
        'We thank the Editor and the two reviewers for their constructive feedback. In this '
        'revision we have (1) added sensitivity and control-series analyses including bootstrap '
        'resampling, post-hoc power simulation, a comparative interrupted time-series model, and '
        'a transaction-level difference-in-differences analysis; (2) softened the causal '
        'interpretation throughout and replaced "ban/prohibited" with accurate wording about a '
        'restriction on routine use with documented clinical exceptions; (3) expanded the '
        'clinical and environmental context, including desflurane pharmacokinetic advantages, '
        'the debate over GWP100 for short-lived anesthetics, and the importance of anesthetic '
        'diversity for supply-chain resilience; and (4) strengthened the limitations section '
        'regarding confounders, eBay generalizability, EU/non-EU transaction uncertainty, and '
        'the small post-restriction sample. A detailed point-by-point response to reviewers is '
        'enclosed.')

    doc.add_paragraph(
        'The manuscript has not been published previously and is not under consideration by '
        'any other journal. All authors have approved the revised manuscript and agree with '
        'its resubmission to the Journal of Anesthesia and Translational Medicine.')

    doc.add_paragraph(
        'We confirm that this study complies with the STROBE guidelines for observational '
        'studies. The completed STROBE checklist is provided as supplementary material. '
        'No ethical approval was required, as the study analyzed publicly available market '
        'data without involving human participants.')

    doc.add_paragraph()

    # Closing
    doc.add_paragraph('Sincerely,')
    doc.add_paragraph()
    doc.add_paragraph('[Corresponding author name]')
    doc.add_paragraph('[Department, Institution]')
    doc.add_paragraph('[Address]')
    doc.add_paragraph('[Email]')

    path = os.path.join(out_dir, 'jatm_cover_letter.docx')
    doc.save(path)
    print(f"JATM cover letter saved: {path}")
    return path


if __name__ == '__main__':
    write_cover_letter()

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

def extract_omml_equations(doc_path):
    doc = Document(doc_path)
    equations = []

    for para in doc.paragraphs:
        # Use qualified namespace for XPath
        omml_elements = para._element.xpath('.//m:oMath')
        for eq in omml_elements:
            equations.append(etree.tostring(eq, encoding='unicode'))

    return equations

OMML_CHAR_WEIGHT = []

def omml_visual_length(omml_xml):
    try:
        tree = etree.fromstring(omml_xml)
        # All text nodes in OMML
        text_nodes = tree.xpath('.//m:t', namespaces={'m': 'http://schemas.openxmlformats.org/officeDocument/2006/math'})
        print(sum(len(t.text or "") for t in text_nodes))
    except Exception:
        return OMML_CHAR_WEIGHT  # fallback



# Example usage
equations = extract_omml_equations('example.docx')

for i, eq in enumerate(equations):
    print(f"Equation {i+1} OMML XML:\n{eq}\n")
    omml_visual_length(eq)
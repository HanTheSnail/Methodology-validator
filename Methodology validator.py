import streamlit as st
import sys
import subprocess
import re

# Fallback: Install dependencies if not available
try:
    from docx import Document
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx", "lxml"])
    from docx import Document

from docx.shared import RGBColor, Pt
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
import io

# Question type requirements - ONLY CRITICAL (upload-blocking) issues
QTYPE_REQUIREMENTS = {
    'Single': {
        'options_required': True,
        'details_required': ['kind', 'select_count'],
        'fix_template': 'kind: exactly\nselect_count: 1'
    },
    'Multi': {
        'options_required': True,
        'details_required': ['kind', 'select_count'],
        'fix_template': 'kind: atleast\nselect_count: 1'
    },
    'Grid': {
        'options_required': True,  # CRITICAL: Upload will fail without options
        'details_required': [],
        'fix_template': 'comment_title: Would you like to add a comment?\ncomment_description: Tell us why'
    },
    'Upload': {
        'options_required': False,
        'details_required': ['type'],  # CRITICAL: Must specify photo/video/audio
        'fix_template': 'type: photo'
    },
    'Image Highlight': {
        'options_required': False,
        'details_required': ['description'],
        'fix_template': 'description: Description\nhint: Answer'
    },
    'AB': {
        'options_required': True,  # CRITICAL: Upload will fail without options
        'details_required': [],
        'fix_template': 'description: Choose A or B'
    },
    'Ranking': {
        'options_required': True,
        'details_required': [],
        'fix_template': 'na_title: Unranked'
    },
    'Open': {
        'options_required': False,
        'details_required': [],
        'fix_template': 'hint: Please explain'
    },
    'Swipe': {
        'options_required': True,
        'details_required': [],
        'fix_template': 'description: Swipe direction'
    },
    'Info Image': {
        'options_required': True,
        'details_required': [],
        'fix_template': ''
    }
}

def extract_cell_text(cell):
    """Extract text from cell including dropdown content controls"""
    from lxml import etree
    
    # Get all text from the cell's XML element
    cell_element = cell._element
    
    # Define namespace
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    # Find all text elements in the cell (including those in content controls)
    text_elements = cell_element.findall('.//w:t', ns)
    
    # Extract all text
    text_parts = [elem.text for elem in text_elements if elem.text]
    
    # Remove duplicates while preserving order
    seen = set()
    unique_parts = []
    for part in text_parts:
        if part not in seen:
            seen.add(part)
            unique_parts.append(part)
    
    result = ' '.join(unique_parts).strip()
    return result if result else ''

import re

def parse_details(details_text):
    """Parse Details column into dict - handles both line-separated and space-separated formats"""
    if not details_text or not details_text.strip():
        return {}
    
    details_dict = {}
    
    # Use regex to find all "key: value" or "key:value" patterns
    # This handles both newline-separated and space-separated formats
    # Pattern: word characters/underscore, optional spaces, colon, optional spaces, value (everything until next key or end)
    pattern = r'(\w+)\s*:\s*([^:\n]+?)(?=\s+\w+\s*:|$)'
    
    matches = re.findall(pattern, details_text, re.DOTALL)
    
    for key, value in matches:
        key = key.strip()
        value = value.strip()
        if key:
            details_dict[key] = value
    
    return details_dict

def validate_question(qtype, qid, content, options, details, row_num):
    """Validate question and return ONLY critical upload-blocking issues"""
    
    # Skip non-question rows
    if not qid or qid in ['QID', ''] or qtype in ['Title', 'Description', 'Group', 'Info']:
        return None
    
    # Skip if Type is empty
    if not qtype or not qtype.strip():
        return {
            'row': row_num,
            'qid': qid,
            'issue': 'Type column is empty',
            'current': {'Type': '', 'Options': options[:50] if options else '', 'Details': details[:50] if details else ''},
            'fix': {'Type': '⚠️ REQUIRED - Cannot determine type from content', 'Options': options, 'Details': details}
        }
    
    # Check if it's a known type
    if qtype not in QTYPE_REQUIREMENTS:
        return None  # Unknown type but not our problem
    
    reqs = QTYPE_REQUIREMENTS[qtype]
    issues = []
    
    # CRITICAL: Check if Options are required but missing
    if reqs['options_required']:
        if not options or not options.strip():
            issues.append(f"Options are REQUIRED for {qtype} questions - upload will FAIL")
    
    # CRITICAL: Check if required Details fields are missing
    details_dict = parse_details(details)
    missing_required = []
    
    for req_field in reqs['details_required']:
        if req_field not in details_dict:
            missing_required.append(req_field)

    
    if missing_required:
        issues.append(f"Details missing required fields: {', '.join(missing_required)} - upload will FAIL")
    
    # If there are critical issues, return the problem with fix
    if issues:
        # Build the fix - only add template if fields are actually missing
        fixed_details = details.strip() if details else ''
        
        if missing_required and reqs['fix_template']:
            # Only add missing fields from template, not the whole template
            template_dict = parse_details(reqs['fix_template'])
            missing_lines = []
            
            for missing_field in missing_required:
                if missing_field in template_dict:
                    missing_lines.append(f"{missing_field}: {template_dict[missing_field]}")
            
            if missing_lines:
                if fixed_details:
                    fixed_details = fixed_details + '\n' + '\n'.join(missing_lines)
                else:
                    fixed_details = '\n'.join(missing_lines)
        
        return {
            'row': row_num,
            'qid': qid,
            'qtype': qtype,
            'issue': ' | '.join(issues),
            'current': {
                'Type': qtype,
                'Options': options[:100] if options else '❌ EMPTY',
                'Details': details[:100] if details else '❌ EMPTY'
            },
            'fix': {
                'Type': qtype,
                'Options': options if options else f'⚠️ ADD OPTIONS HERE (2+ choices for {qtype})',
                'Details': fixed_details
            }
        }
    
    return None

def analyze_document(doc):
    """Analyze document and return ONLY critical issues"""
    critical_issues = []
    
    # Find the methodology table
    table = None
    for t in doc.tables:
        if len(t.rows) > 0:
            # Try to extract text including dropdowns
            headers = [extract_cell_text(cell) for cell in t.rows[0].cells]
            if 'Type' in headers and 'QID' in headers:
                table = t
                break
    
    if not table:
        return None, "No valid methodology table found"
    
    # Validate each row
    for idx, row in enumerate(table.rows[1:], start=2):
        cells = [extract_cell_text(cell) for cell in row.cells]
        
        if len(cells) >= 5:
            qtype, qid, content, options, details = cells[0], cells[1], cells[2], cells[3], cells[4]
            
            issue = validate_question(qtype, qid, content, options, details, idx)
            if issue:
                critical_issues.append(issue)
    
    return critical_issues, None

# Streamlit UI
st.set_page_config(page_title="Methodology Validator", page_icon="🔍", layout="wide")

st.title("🔍 Questionnaire Methodology Validator")
st.markdown("### Upload your methodology document to check for **critical upload-blocking issues**")

st.info("ℹ️ This tool only flags issues that will **prevent your questionnaire from uploading**. Optional formatting suggestions are not shown.")

uploaded_file = st.file_uploader("Upload Methodology Document (.docx)", type=['docx'])

if uploaded_file:
    doc = Document(uploaded_file)
    
    st.success("✅ Document loaded successfully!")
    
    with st.spinner("Analyzing for critical issues..."):
        critical_issues, error = analyze_document(doc)
    
    if error:
        st.error(error)
    elif not critical_issues:
        st.balloons()
        st.success("🎉 **No critical issues found!** Your methodology is ready to upload.")
    else:
        st.error(f"⚠️ Found **{len(critical_issues)} critical issues** that will prevent upload")
        
        st.markdown("---")
        st.markdown("## 🔧 Issues & Fixes")
        st.markdown("Here's what needs to be fixed in your document:")
        
        for i, issue in enumerate(critical_issues, 1):
            st.markdown(f"### Issue {i}: Question {issue['qid']} (Row {issue['row']})")
            
            # Show the problem clearly
            st.error(f"**Problem:** {issue['issue']}")
            
            # Show current state vs fixed state
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**❌ Current (Broken):**")
                st.code(f"""Type: {issue['current']['Type']}
QID: {issue['qid']}
Options: {issue['current']['Options']}
Details: {issue['current']['Details']}""", language=None)
            
            with col2:
                st.markdown("**✅ Fixed (Copy This):**")
                st.code(f"""Type: {issue['fix']['Type']}
QID: {issue['qid']}
Options: {issue['fix']['Options']}
Details: {issue['fix']['Details']}""", language=None)
            
            st.markdown("---")
        
        # Summary
        st.markdown("### 📋 Summary")
        st.warning(f"""
**Action Required:** Fix {len(critical_issues)} row(s) in your Word document before uploading.

**How to fix:**
1. Open your methodology Word document
2. For each issue above, update the row to match the "✅ Fixed" version
3. Re-upload here to verify all issues are resolved
        """)

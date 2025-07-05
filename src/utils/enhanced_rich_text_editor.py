"""
Enhanced Rich Text Editor for NSNA Mail Merge Tool
Provides sophisticated editing features including:
- Bold, Italic formatting
- Bullet lists and numbered lists
- Line spacing controls
- Variable insertion buttons
- Real-time preview
"""

import streamlit as st
import re
from typing import List, Dict


def create_enhanced_rich_text_editor(
    value: str = "",
    placeholder: str = "Enter your content here...",
    key: str = "enhanced_editor",
    height: int = 300,
    available_variables: List[str] = None,
    show_formatting_help: bool = True
) -> str:
    """
    Create an enhanced rich text editor with formatting tools and variable insertion
    
    Args:
        value: Initial content
        placeholder: Placeholder text
        key: Unique key for the editor
        height: Editor height in pixels
        available_variables: List of variables that can be inserted
        show_formatting_help: Whether to show formatting help
    
    Returns:
        The edited content
    """
    
    if available_variables is None:
        available_variables = ["First Name", "Last Name", "Amount", "Date", "Year"]
    
    # Create a container for the editor
    editor_container = st.container()
    
    with editor_container:
        # Formatting toolbar
        st.markdown("### ✏️ Rich Text Editor")
        
        # Create columns for toolbar
        col1, col2, col3, col4 = st.columns([2, 2, 2, 4])
        
        with col1:
            st.markdown("**Format:**")
            if st.button("**B** Bold", key=f"{key}_bold", help="Wrap selected text with **bold**"):
                st.session_state[f"{key}_insert"] = "**bold text**"
            if st.button("*I* Italic", key=f"{key}_italic", help="Wrap selected text with *italic*"):
                st.session_state[f"{key}_insert"] = "*italic text*"
        
        with col2:
            st.markdown("**Lists:**")
            if st.button("• Bullet", key=f"{key}_bullet", help="Insert bullet point"):
                st.session_state[f"{key}_insert"] = "• Bullet point"
            if st.button("1. Number", key=f"{key}_number", help="Insert numbered list"):
                st.session_state[f"{key}_insert"] = "1. Numbered item"
        
        with col3:
            st.markdown("**Spacing:**")
            if st.button("¶ Paragraph", key=f"{key}_paragraph", help="Add paragraph break"):
                st.session_state[f"{key}_insert"] = "\n\n"
            if st.button("↵ Line Break", key=f"{key}_linebreak", help="Add line break"):
                st.session_state[f"{key}_insert"] = "\n"
        
        with col4:
            st.markdown("**Variables:**")
            # Variable insertion dropdown
            selected_var = st.selectbox(
                "Insert Variable",
                [""] + available_variables,
                key=f"{key}_var_select",
                help="Select a variable to insert"
            )
            
            if selected_var:
                if st.button(f"Insert {{{selected_var}}}", key=f"{key}_insert_var"):
                    st.session_state[f"{key}_insert"] = f"{{{selected_var}}}"
                    # Reset the selectbox
                    st.session_state[f"{key}_var_select"] = ""
                    st.rerun()
        
        # Check if there's content to insert
        insert_content = st.session_state.get(f"{key}_insert", "")
        if insert_content:
            # Add the content to the current value
            cursor_pos = len(value)  # Simple insertion at the end
            value = value + insert_content
            # Clear the insert flag
            st.session_state[f"{key}_insert"] = ""
        
        # Main text editor
        content = st.text_area(
            "Content",
            value=value,
            placeholder=placeholder,
            key=f"{key}_main",
            height=height,
            help="Use the toolbar above to format your text, or type markdown directly"
        )
        
        # Formatting help section
        if show_formatting_help:
            with st.expander("📖 Formatting Help"):
                st.markdown("""
                **Markdown Formatting Guide:**
                
                - `**bold text**` → **bold text**
                - `*italic text*` → *italic text*
                - `• bullet point` → • bullet point
                - `1. numbered item` → 1. numbered item
                
                **Variables:**
                
                Use curly braces to insert variables: `{First Name}`, `{Last Name}`, etc.
                
                **Line Breaks:**
                
                - Single line break: Press Enter once
                - Paragraph break: Press Enter twice (empty line)
                
                **Tips:**
                
                - Use the toolbar buttons for quick formatting
                - Select variables from the dropdown to insert them
                - Preview your content in the preview section below
                """)
        
        # Live preview section
        if content and content.strip():
            st.markdown("---")
            st.markdown("### 👀 Live Preview")
            
            # Convert markdown to HTML for preview
            preview_content = content
            preview_content = re.sub(r'\*\*([^*]+)\*\*', r'<strong>\1</strong>', preview_content)
            preview_content = re.sub(r'(?<!\*)\*([^*]+)\*(?!\*)', r'<em>\1</em>', preview_content)
            
            # Convert line breaks to HTML
            preview_content = preview_content.replace('\n\n', '<br><br>')
            preview_content = preview_content.replace('\n', '<br>')
            
            # Display preview
            st.markdown(preview_content, unsafe_allow_html=True)
            
            # Show character count
            st.caption(f"📊 Character count: {len(content)} | Words: {len(content.split())}")
    
    return content


def create_variable_reference_sidebar(available_variables: List[str]):
    """Create a sidebar with variable reference"""
    
    with st.sidebar:
        st.markdown("### 📋 Available Variables")
        st.markdown("Click to copy:")
        
        for var in available_variables:
            if st.button(f"{{{var}}}", key=f"sidebar_var_{var}"):
                st.code(f"{{{var}}}")
                st.success(f"Variable {{{var}}} ready to copy!")
        
        st.markdown("---")
        st.markdown("""
        **Quick Formatting:**
        - **Bold**: `**text**`
        - *Italic*: `*text*`
        - Bullet: `• item`
        - Number: `1. item`
        """)


def enhanced_rich_text_editor_with_templates(
    value: str = "",
    key: str = "enhanced_editor",
    available_variables: List[str] = None,
    template_type: str = "email"  # "email" or "pdf"
) -> str:
    """
    Enhanced rich text editor with pre-built templates
    """
    
    if available_variables is None:
        available_variables = ["First Name", "Last Name", "Amount", "Date", "Year"]
    
    # Template selection
    templates = {
        "Blank": "",
        "Thank You Letter": f"""Dear {{First Name}} {{Last Name}},

**Thank you** for your generous donation to our organization. This letter will serve as a tax receipt for your contribution.

**Donation Details:**
• Amount: ${{Amount}}
• Date: {{Date}}
• Year: {{Year}}

*No goods or services were provided in exchange for your contribution.*

We truly appreciate your support!

**Sincerely,**
NSNA Team""",
        
        "Receipt Template": f"""**DONATION RECEIPT**

Dear {{First Name}} {{Last Name}},

On behalf of Nagarathar Sangam of North America, thank you for your recent donation to our organization.

**Donation Information:**
• Donor: {{First Name}} {{Last Name}}
• Amount: ${{Amount}}
• Date: {{Date}}
• Tax Year: {{Year}}

This letter serves as your **official tax receipt** for your contribution listed above.

*Nagarathar Sangam of North America is a registered Section 501(c)(3) non-profit organization.*

Thank you for your continued support of our mission.

**Sincerely,**
Treasurer, NSNA
({{Year}}-{int(st.session_state.get('current_year', 2025))+1} term)""",
        
        "Simple Acknowledgment": f"""Dear {{First Name}},

Thank you for your donation of ${{Amount}} received on {{Date}}.

Your contribution helps support our mission and community programs.

*This serves as your tax receipt for {{Year}}.*

Best regards,
NSNA Team"""
    }
    
    # Template selector
    col1, col2 = st.columns([3, 1])
    
    with col1:
        selected_template = st.selectbox(
            "Choose a template:",
            list(templates.keys()),
            key=f"{key}_template_select"
        )
    
    with col2:
        if st.button("Load Template", key=f"{key}_load_template"):
            value = templates[selected_template]
            st.session_state[f"{key}_template_loaded"] = True
            st.rerun()
    
    # Use template content if loaded
    if st.session_state.get(f"{key}_template_loaded", False):
        value = templates[selected_template]
        st.session_state[f"{key}_template_loaded"] = False
    
    # Enhanced editor
    content = create_enhanced_rich_text_editor(
        value=value,
        key=key,
        available_variables=available_variables,
        height=400
    )
    
    return content

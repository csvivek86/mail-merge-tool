"""
Rich text editor component for NSNA Mail Merge Tool
Provides Word-like email template editing with intuitive formatting
"""

import streamlit as st
from datetime import datetime
from typing import List, Dict, Any
import re

def create_rich_text_editor(content: str, available_variables: List[str], key: str = "rich_editor") -> str:
    """
    Create a Word-like rich text editor with intuitive formatting and proper variable insertion
    """
    
    # Initialize content in session state if not exists
    if f"{key}_content" not in st.session_state:
        st.session_state[f"{key}_content"] = content
    
    st.markdown("**✏️ Email Content Editor**")
    st.markdown("Create your email content using the formatting tools below, just like in Microsoft Word!")
    
    # Word-like formatting toolbar
    st.markdown("**🛠️ Formatting Toolbar**")
    
    # Row 1: Text formatting
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        if st.button("**B** Bold", help="Make selected text bold", key=f"{key}_bold"):
            st.session_state[f"{key}_format_action"] = "bold"
    
    with col2:
        if st.button("*I* Italic", help="Make selected text italic", key=f"{key}_italic"):
            st.session_state[f"{key}_format_action"] = "italic"
    
    with col3:
        if st.button("U̲ Underline", help="Underline selected text", key=f"{key}_underline"):
            st.session_state[f"{key}_format_action"] = "underline"
    
    with col4:
        if st.button("• Bullets", help="Create bullet list", key=f"{key}_bullets"):
            st.session_state[f"{key}_format_action"] = "bullets"
    
    with col5:
        if st.button("🔢 Numbers", help="Create numbered list", key=f"{key}_numbers"):
            st.session_state[f"{key}_format_action"] = "numbers"
    
    # Row 2: Content insertion
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        if st.button("📅 Today's Date", help="Insert current date", key=f"{key}_date"):
            current_date = datetime.now().strftime("%B %d, %Y")
            st.session_state[f"{key}_insert_text"] = current_date
    
    with col2:
        if st.button("📆 Current Year", help="Insert current year", key=f"{key}_year"):
            current_year = str(datetime.now().year)
            st.session_state[f"{key}_insert_text"] = current_year
    
    with col3:
        # Variable dropdown with better organization
        if available_variables:
            # Create organized variable list
            var_options = ["Click to insert variable..."]
            
            # Add commonly used variables first
            common_vars = []
            financial_vars = []
            other_vars = []
            
            for var in available_variables:
                var_lower = var.lower()
                if any(keyword in var_lower for keyword in ['first', 'last', 'name', 'email']):
                    common_vars.append(var)
                elif any(keyword in var_lower for keyword in ['amount', 'donation', 'total', 'sum', 'payment']):
                    financial_vars.append(var)
                else:
                    other_vars.append(var)
            
            if common_vars:
                var_options.append("--- Personal Info ---")
                var_options.extend(common_vars)
            
            if financial_vars:
                var_options.append("--- Financial ---")
                var_options.extend(financial_vars)
            
            if other_vars:
                var_options.append("--- Other ---")
                var_options.extend(other_vars)
            
            selected_var = st.selectbox(
                "Insert Variable",
                var_options,
                key=f"{key}_var_select",
                help="Select a variable to insert from your Excel data"
            )
            
            if selected_var and not selected_var.startswith("---") and selected_var != "Click to insert variable...":
                st.session_state[f"{key}_insert_text"] = f"{{{selected_var}}}"
    
    with col4:
        if st.button("�️ Clear All", help="Clear all content", key=f"{key}_clear"):
            st.session_state[f"{key}_content"] = ""
            st.session_state[f"{key}_clear_flag"] = True
    
    # Content insertion instructions
    st.markdown("""
    **💡 How to use:**
    - Type your message in the text area below
    - Select text and use formatting buttons to apply styles
    - Click variable buttons to insert at cursor position
    - Use the preview to see how it will look
    """)
    
    # Handle formatting actions
    current_content = st.session_state[f"{key}_content"]
    
    # Handle text insertion
    if f"{key}_insert_text" in st.session_state:
        insert_text = st.session_state[f"{key}_insert_text"]
        # For now, append to end (cursor position detection is limited in Streamlit)
        current_content = current_content + " " + insert_text if current_content else insert_text
        st.session_state[f"{key}_content"] = current_content
        del st.session_state[f"{key}_insert_text"]
    
    # Handle formatting actions
    if f"{key}_format_action" in st.session_state:
        action = st.session_state[f"{key}_format_action"]
        if action == "bold":
            st.info("💡 Tip: Type **your bold text** to make it bold")
        elif action == "italic":
            st.info("💡 Tip: Type *your italic text* to make it italic")
        elif action == "underline":
            st.info("💡 Tip: Type <u>your underlined text</u> to underline")
        elif action == "bullets":
            if current_content and not current_content.endswith('\n'):
                current_content += '\n'
            current_content += "• "
            st.session_state[f"{key}_content"] = current_content
        elif action == "numbers":
            if current_content and not current_content.endswith('\n'):
                current_content += '\n'
            current_content += "1. "
            st.session_state[f"{key}_content"] = current_content
        
        del st.session_state[f"{key}_format_action"]
    
    # Clear flag handling
    if f"{key}_clear_flag" in st.session_state:
        st.session_state[f"{key}_content"] = ""
        del st.session_state[f"{key}_clear_flag"]
        current_content = ""
    
    # Main text area with larger size
    edited_content = st.text_area(
        "Email Content",
        value=current_content,
        height=400,
        help="Type your email content here. Use the toolbar above for formatting and variable insertion.",
        key=f"{key}_textarea",
        placeholder="Start typing your email content here...\n\nUse the formatting buttons above to style your text.\nClick 'Insert Variable' to add recipient data.\nClick 'Today's Date' or 'Current Year' to add current information."
    )
    
    # Update session state
    st.session_state[f"{key}_content"] = edited_content
    
    # Live preview section
    if edited_content.strip():
        st.markdown("---")
        st.markdown("**👀 Live Preview**")
        
        # Convert basic markdown-style formatting to HTML-like preview
        preview_content = edited_content
        
        # Apply basic formatting
        preview_content = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', preview_content)
        preview_content = re.sub(r'\*(.*?)\*', r'<em>\1</em>', preview_content)
        preview_content = re.sub(r'<u>(.*?)</u>', r'<u>\1</u>', preview_content)
        
        # Handle bullet points
        preview_content = re.sub(r'^• (.+)$', r'<li>\1</li>', preview_content, flags=re.MULTILINE)
        preview_content = re.sub(r'^- (.+)$', r'<li>\1</li>', preview_content, flags=re.MULTILINE)
        
        # Handle numbered lists
        preview_content = re.sub(r'^\d+\. (.+)$', r'<li>\1</li>', preview_content, flags=re.MULTILINE)
        
        # Wrap consecutive list items in ul tags
        preview_content = re.sub(r'(<li>.*?</li>(\s*<li>.*?</li>)*)', r'<ul>\1</ul>', preview_content, flags=re.DOTALL)
        
        # Show preview in a styled container
        st.markdown(f"""
        <div style="border: 1px solid #ddd; padding: 20px; border-radius: 8px; background-color: #f8f9fa; margin: 10px 0; font-family: Arial, sans-serif; line-height: 1.6;">
            {preview_content.replace(chr(10), '<br>')}
        </div>
        """, unsafe_allow_html=True)
        
        # Variable validation
        variables_used = re.findall(r'\{([^}]+)\}', edited_content)
        if variables_used:
            st.markdown("**🔍 Variables Used:**")
            valid_vars = []
            invalid_vars = []
            
            for var in set(variables_used):
                if var in available_variables:
                    valid_vars.append(var)
                else:
                    invalid_vars.append(var)
            
            if valid_vars:
                st.success(f"✅ Valid variables: {', '.join(valid_vars)}")
            
            if invalid_vars:
                st.error(f"❌ Invalid variables: {', '.join(invalid_vars)}")
                st.warning("Make sure variable names exactly match your Excel column headers!")
    
    return edited_content

def create_template_formatting_options() -> Dict[str, Any]:
    """
    Create advanced formatting options for email templates
    """
    st.markdown("**🎨 Advanced Formatting Options**")
    
    col1, col2 = st.columns(2)
    
    with col1:
        font_family = st.selectbox(
            "Font Family",
            ["Arial", "Georgia", "Times New Roman", "Helvetica", "Verdana"],
            index=0,
            help="Choose the font family for your email"
        )
        
        font_size = st.selectbox(
            "Font Size",
            ["12px", "14px", "16px", "18px", "20px"],
            index=1,
            help="Choose the font size for your email"
        )
    
    with col2:
        text_color = st.selectbox(
            "Text Color",
            ["#333333 (Dark Gray)", "#000000 (Black)", "#2E4BC6 (Blue)", "#C62E2E (Red)", "#2E8B57 (Green)"],
            index=0,
            help="Choose the text color for your email"
        )
        
        line_spacing = st.selectbox(
            "Line Spacing",
            ["1.4", "1.6", "1.8", "2.0"],
            index=1,
            help="Choose the line spacing for your email"
        )
    
    # Extract color code
    color_code = text_color.split(" ")[0]
    
    return {
        'font_family': font_family,
        'font_size': font_size,
        'text_color': color_code,
        'line_spacing': line_spacing
    }

def apply_formatting_to_html(content: str, formatting: Dict[str, Any]) -> str:
    """
    Apply formatting options to HTML content
    """
    # Convert markdown-style formatting to HTML
    html_content = content
    
    # Apply basic formatting
    html_content = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', html_content)
    html_content = re.sub(r'\*(.*?)\*', r'<em>\1</em>', html_content)
    html_content = re.sub(r'<u>(.*?)</u>', r'<u>\1</u>', html_content)
    
    # Handle bullet points
    html_content = re.sub(r'^• (.+)$', r'<li>\1</li>', html_content, flags=re.MULTILINE)
    html_content = re.sub(r'^- (.+)$', r'<li>\1</li>', html_content, flags=re.MULTILINE)
    
    # Handle numbered lists
    html_content = re.sub(r'^\d+\. (.+)$', r'<li>\1</li>', html_content, flags=re.MULTILINE)
    
    # Wrap consecutive list items in ul tags
    html_content = re.sub(r'(<li>.*?</li>(\s*<li>.*?</li>)*)', r'<ul>\1</ul>', html_content, flags=re.DOTALL)
    
    # Create full HTML with styling
    formatted_html = f"""
    <html>
    <body style="
        font-family: {formatting['font_family']}, sans-serif; 
        line-height: {formatting['line_spacing']}; 
        color: {formatting['text_color']};
        font-size: {formatting['font_size']};
        max-width: 600px;
        margin: 0 auto;
        padding: 20px;
    ">
        <div>
            {html_content.replace(chr(10), '<br>')}
        </div>
    </body>
    </html>
    """
    return formatted_html

def create_variable_helper(available_variables: List[str]) -> None:
    """
    Create a helpful display of available variables organized by category
    """
    if not available_variables:
        st.info("📋 No variables available. Upload Excel data first to see available variables.")
        return
    
    st.markdown("**📋 Available Variables from Your Excel Data**")
    
    # Organize variables by category
    personal_vars = []
    financial_vars = []
    date_vars = []
    other_vars = []
    
    for var in available_variables:
        var_lower = var.lower()
        if any(keyword in var_lower for keyword in ['name', 'first', 'last', 'email', 'phone']):
            personal_vars.append(var)
        elif any(keyword in var_lower for keyword in ['amount', 'donation', 'total', 'sum', 'payment', 'price', 'cost']):
            financial_vars.append(var)
        elif any(keyword in var_lower for keyword in ['date', 'year', 'time', 'month', 'day']):
            date_vars.append(var)
        else:
            other_vars.append(var)
    
    # Display in organized columns
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        if personal_vars:
            st.markdown("**👤 Personal Info:**")
            for var in personal_vars:
                st.code(f"{{{var}}}", language=None)
    
    with col2:
        if financial_vars:
            st.markdown("**💰 Financial:**")
            for var in financial_vars:
                st.code(f"{{{var}}}", language=None)
    
    with col3:
        if date_vars:
            st.markdown("**📅 Dates:**")
            for var in date_vars:
                st.code(f"{{{var}}}", language=None)
    
    with col4:
        if other_vars:
            st.markdown("**📊 Other:**")
            for var in other_vars[:8]:  # Limit to 8 to avoid clutter
                st.code(f"{{{var}}}", language=None)
            if len(other_vars) > 8:
                st.caption(f"... and {len(other_vars) - 8} more")

def create_email_template_presets() -> Dict[str, str]:
    """
    Create preset email templates for common use cases
    """
    presets = {
        "Donation Receipt": """Dear {First Name} {Last Name},

**Thank you for your generous donation to the Nagarathar Sangam of North America.**

We are pleased to acknowledge your contribution of **${Amount}** received on **{Date}**.

Your support helps us:
• Preserve our cultural heritage
• Support community programs  
• Maintain our traditional values
• Strengthen our community bonds

This letter serves as your official receipt for tax purposes.

Best regards,

NSNA Treasurer
Nagarathar Sangam of North America
Email: treasurer@nsna.org""",
        
        "Event Thank You": """Dear {First Name} {Last Name},

**Thank you for attending our recent NSNA event!**

Your participation and support mean the world to our community.

We hope you enjoyed:
• The cultural presentations
• Networking with fellow members
• The delicious traditional food
• The meaningful conversations

We look forward to seeing you at future gatherings.

With warm regards,

NSNA Event Committee""",
        
        "Membership Welcome": """Dear {First Name} {Last Name},

**Welcome to the Nagarathar Sangam of North America!**

We are delighted to have you as a member of our vibrant community.

As a member, you'll enjoy:
• Access to all NSNA events
• Cultural preservation programs
• Networking opportunities
• Community support services

Your membership helps us preserve our traditions and strengthen our bonds as a community.

We look forward to your active participation in our activities and events.

Best regards,

NSNA Membership Committee""",
        
        "Custom Template": ""
    }
    
    return presets
    
    col1, col2 = st.columns(2)
    
    with col1:
        font_family = st.selectbox(
            "Font Family",
            ["Arial", "Georgia", "Times New Roman", "Helvetica", "Verdana"],
            index=0
        )
        
        font_size = st.selectbox(
            "Font Size",
            ["12px", "14px", "16px", "18px", "20px"],
            index=1
        )
    
    with col2:
        text_color = st.selectbox(
            "Text Color",
            ["#333333 (Dark Gray)", "#000000 (Black)", "#2E4BC6 (Blue)", "#C62E2E (Red)", "#2E8B57 (Green)"],
            index=0
        )
        
        line_spacing = st.selectbox(
            "Line Spacing",
            ["1.4", "1.6", "1.8", "2.0"],
            index=1
        )
    
    # Extract color code
    color_code = text_color.split(" ")[0]
    
    return {
        'font_family': font_family,
        'font_size': font_size,
        'text_color': color_code,
        'line_spacing': line_spacing
    }

def apply_formatting_to_html(content: str, formatting: Dict[str, Any]) -> str:
    """
    Apply formatting options to HTML content
    """
    html_content = f"""
    <html>
    <body style="
        font-family: {formatting['font_family']}, sans-serif; 
        line-height: {formatting['line_spacing']}; 
        color: {formatting['text_color']};
        font-size: {formatting['font_size']};
    ">
        <div style="max-width: 600px; margin: 0 auto; padding: 20px;">
            {content.replace(chr(10), '<br>')}
        </div>
    </body>
    </html>
    """
    return html_content

def create_variable_helper(available_variables: List[str]) -> None:
    """
    Create a helper section showing available variables
    """
    if not available_variables:
        return
    
    st.markdown("**Available Variables** 📋")
    
    # Group variables by type for better organization
    personal_vars = [var for var in available_variables if any(keyword in var.lower() for keyword in ['name', 'first', 'last', 'email'])]
    financial_vars = [var for var in available_variables if any(keyword in var.lower() for keyword in ['amount', 'donation', 'total', 'sum'])]
    date_vars = [var for var in available_variables if any(keyword in var.lower() for keyword in ['date', 'year', 'time'])]
    other_vars = [var for var in available_variables if var not in personal_vars + financial_vars + date_vars]
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        if personal_vars:
            st.markdown("**👤 Personal:**")
            for var in personal_vars[:5]:  # Limit to 5 items
                st.code(f"{{{var}}}")
    
    with col2:
        if financial_vars:
            st.markdown("**💰 Financial:**")
            for var in financial_vars[:5]:
                st.code(f"{{{var}}}")
    
    with col3:
        if date_vars:
            st.markdown("**📅 Dates:**")
            for var in date_vars[:5]:
                st.code(f"{{{var}}}")
    
    with col4:
        if other_vars:
            st.markdown("**📊 Other:**")
            for var in other_vars[:5]:
                st.code(f"{{{var}}}")

def create_email_template_presets() -> Dict[str, str]:
    """
    Create preset email templates for common use cases
    """
    presets = {
        "Donation Receipt": """Dear {First Name} {Last Name},

Thank you for your generous donation to the Nagarathar Sangam of North America.

We are pleased to acknowledge your contribution of **${Amount}** received on **{Date}**.

Your support helps us continue our mission of preserving our cultural heritage and supporting our community.

This letter serves as your official receipt for tax purposes.

Best regards,

NSNA Treasurer
Nagarathar Sangam of North America""",
        
        "Event Thank You": """Dear {First Name} {Last Name},

Thank you for attending our recent NSNA event!

Your participation and support mean the world to our community.

We hope you enjoyed the event and look forward to seeing you at future gatherings.

Best regards,

NSNA Event Committee""",
        
        "Membership Welcome": """Dear {First Name} {Last Name},

Welcome to the Nagarathar Sangam of North America!

We are delighted to have you as a member of our community.

Your membership helps us preserve our cultural traditions and strengthen our bonds as a community.

We look forward to your participation in our activities and events.

Best regards,

NSNA Membership Committee""",
        
        "Custom Template": ""
    }
    
    return presets

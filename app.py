import streamlit as st
import pandas as pd
import os
import time
import json
import re
import io
import base64
from lead_engine_generator import LeadEngineGenerator

# Set page configuration
st.set_page_config(
    page_title="Lead Engine Content Generator",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1E3A8A;
        margin-bottom: 1rem;
        text-align: center;
    }
    .sub-header {
        font-size: 1.5rem;
        font-weight: bold;
        color: #1E3A8A;
        margin-top: 2rem;
        margin-bottom: 1rem;
    }
    .stProgress .st-eb {
        background-color: #1E3A8A;
    }
    .success-message {
        background-color: #ECFDF5;
        color: #047857;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
    .info-message {
        background-color: #EFF6FF;
        color: #1E40AF;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
    .instruction-message {
        background-color: #F3F4F6;
        color: #374151;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
        border-left: 4px solid #1E3A8A;
    }
    .feature-box {
        background-color: #F9FAFB;
        border-radius: 0.5rem;
        padding: 1rem;
        margin-bottom: 1rem;
        border: 1px solid #E5E7EB;
    }
</style>
""", unsafe_allow_html=True)

# Title and introduction - Always visible at the top
st.markdown('<div class="main-header">✨ Lead Engine Content Generator ✨</div>', unsafe_allow_html=True)
st.markdown("""
Generate a complete social media marketing funnel for your client with AI-powered content.
This tool automatically creates brand, demand generation, and demand capture content.
""")

# Initialize session state variables if they don't exist
if 'generator' not in st.session_state:
    st.session_state.generator = None
if 'content_generated' not in st.session_state:
    st.session_state.content_generated = False
if 'excel_data' not in st.session_state:
    st.session_state.excel_data = None
if 'funnel_content' not in st.session_state:
    st.session_state.funnel_content = {}
if 'processing_time' not in st.session_state:
    st.session_state.processing_time = 0

# Create sidebar for inputs
with st.sidebar:
    st.markdown('<div class="sub-header">Configuration</div>', unsafe_allow_html=True)
    
    # OpenAI API Key input
    api_key = st.text_input("OpenAI API Key", type="password", help="Your OpenAI API key is required to generate content")
    
    # Website URL input
    website_url = st.text_input("Client's Website URL", help="Enter the full URL including http:// or https://")
    
    # Lead purpose selection
    lead_purpose = st.selectbox(
        "Lead Purpose",
        ["Demo Booking", "Sales Meeting"],
        help="Select the primary goal for lead generation"
    )
    
    # Asset link input with conditional display
    has_asset = st.checkbox("Include downloadable asset", value=True)
    asset_link = None
    if has_asset:
        asset_link = st.text_input("Asset Link", help="URL where leads can download your asset")
        
        # PDF URL for content extraction (optional)
        pdf_url = st.text_input(
            "PDF URL for Content Extraction (Optional)", 
            help="Link to a PDF that contains relevant content for your marketing"
        )
    else:
        pdf_url = None
    
    # Number of posts selection
    num_posts = st.slider(
        "Number of posts per funnel stage", 
        min_value=1,
        max_value=10,
        value=3,
        help="How many posts to generate for each marketing funnel stage"
    )
    
    # Channel selection
    channel_options = st.multiselect(
        "Social Media Channels",
        ["LinkedIn", "Facebook"],
        default=["LinkedIn", "Facebook"],
        help="Select which social media platforms to generate content for"
    )
    
    # Generate button
    generate_button = st.button("Generate Content Funnel", type="primary", use_container_width=True)

# Function to generate content
def generate_content():
    # Display progress information
    progress_container = st.container()
    progress_bar = progress_container.progress(0)
    status_text = progress_container.empty()
    
    try:
        # Initialize the generator
        status_text.text("Initializing Lead Engine Generator...")
        generator = LeadEngineGenerator(api_key)
        st.session_state.generator = generator
        progress_bar.progress(10)
        
        # Set lead purpose
        status_text.text(f"Setting lead purpose to: {lead_purpose}")
        generator.set_lead_purpose(lead_purpose, asset_link)
        progress_bar.progress(20)
        
        # Start timing
        start_time = time.time()
        
        # Generate content
        status_text.text("Generating content funnel... This may take a few minutes.")
        success = generator.generate_funnel_content(website_url, pdf_url, num_posts, channel_options)
        progress_bar.progress(90)
        
        if success:
            # Store the funnel content in session state
            st.session_state.funnel_content = generator.funnel_content
            st.session_state.content_generated = True
            
            # Create an in-memory Excel file
            status_text.text("Preparing Excel file for download...")
            excel_data = generate_excel_file(generator)
            st.session_state.excel_data = excel_data
            
            # Calculate processing time
            st.session_state.processing_time = time.time() - start_time
            
            progress_bar.progress(100)
            status_text.text("Content generation completed!")
            
            # Return success
            return True
        else:
            status_text.error("Failed to generate content funnel.")
            progress_bar.progress(100)
            return False
            
    except Exception as e:
        status_text.error(f"Error: {str(e)}")
        progress_bar.progress(100)
        return False

# Function to generate Excel file in memory
def generate_excel_file(generator):
    # Create BytesIO object to hold Excel data
    output = io.BytesIO()
    
    # Create pandas ExcelWriter
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Process each channel
        for channel in generator.funnel_content.keys():
            # Create DataFrame for channel sheet
            channel_data = []
            
            # Add Brand posts
            if "Brand" in generator.funnel_content[channel]:
                # Add section header row
                channel_data.append({
                    "Post #": "",
                    "Value Proposition": "",
                    "Post Copy": "",
                    "Image Copy": "",
                    "CTA": "",
                    "Link": "",
                    "Type": "Brand"
                })
                
                # Better error handling for posts
                brand_posts = generator.funnel_content[channel]["Brand"]
                if brand_posts and isinstance(brand_posts, list):
                    for i, post in enumerate(brand_posts):
                        if not isinstance(post, dict):
                            continue
                        
                        row = {
                            "Post #": i + 1,
                            "Value Proposition": post.get("Value Proposition", ""),
                            "Post Copy": post.get("Post Copy", ""),
                            "Image Copy": post.get("Image Copy", ""),
                            "CTA": "",
                            "Link": "",
                            "Type": ""
                        }
                        channel_data.append(row)
           
            # Add Demand Gen posts
            if "Demand Gen" in generator.funnel_content[channel]:
                # Add section header row
                channel_data.append({
                    "Post #": "",
                    "Value Proposition": "",
                    "Post Copy": "",
                    "Image Copy": "",
                    "CTA": "",
                    "Link": "",
                    "Type": "Demand Gen"
                })
                
                demand_gen_posts = generator.funnel_content[channel]["Demand Gen"]
                if demand_gen_posts and isinstance(demand_gen_posts, list):
                    for i, post in enumerate(demand_gen_posts):
                        if not isinstance(post, dict):
                            continue
                            
                        post_type = post.get("Type", "")
                        row = {
                            "Post #": i + 1,
                            "Value Proposition": "",
                            "Post Copy": post.get("Post Copy", ""),
                            "Image Copy": post.get("Image Copy", ""),
                            "CTA": post.get("CTA", ""),
                            "Link": post.get("Link", ""),
                            "Type": post_type
                        }
                        channel_data.append(row)
            
            # Add Demand Capture posts
            if "Demand Capture" in generator.funnel_content[channel]:
                # Add section header row
                channel_data.append({
                    "Post #": "",
                    "Value Proposition": "",
                    "Post Copy": "",
                    "Image Copy": "",
                    "CTA": "",
                    "Link": "",
                    "Type": "Demand Capture"
                })
                
                demand_capture_posts = generator.funnel_content[channel]["Demand Capture"]
                if demand_capture_posts and isinstance(demand_capture_posts, list):
                    for i, post in enumerate(demand_capture_posts):
                        if not isinstance(post, dict):
                            continue
                            
                        row = {
                            "Post #": i + 1,
                            "Value Proposition": "",
                            "Post Copy": post.get("Post Copy", ""),
                            "Image Copy": post.get("Image Copy", ""),
                            "CTA": post.get("CTA", ""),
                            "Link": post.get("Link", ""),
                            "Type": generator.lead_purpose if generator.lead_purpose else "Demand Capture"
                        }
                        channel_data.append(row)
            
            # Create DataFrame and export to Excel
            df = pd.DataFrame(channel_data)
            df.to_excel(writer, sheet_name=channel, index=False)
            
            # Apply styling (simplified for in-memory version)
            worksheet = writer.sheets[channel]
            for col_num, column_title in enumerate(df.columns, 1):
                if column_title in ["Post Copy", "Value Proposition"]:
                    worksheet.column_dimensions[get_column_letter(col_num)].width = 40
                elif column_title in ["Image Copy", "CTA", "Link"]:
                    worksheet.column_dimensions[get_column_letter(col_num)].width = 25
                else:
                    worksheet.column_dimensions[get_column_letter(col_num)].width = 15
    
    # Get the Excel data
    output.seek(0)
    return output.getvalue()

# Helper function for column letter (since we're not importing openpyxl directly)
def get_column_letter(col_num):
    letters = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters

# Create download link for Excel file
def get_excel_download_link(excel_data, filename="marketing_funnel.xlsx"):
    b64 = base64.b64encode(excel_data).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}" class="download-button">Download Excel File</a>'
    return href

# Display content preview
def display_content_preview():
    st.markdown('<div class="sub-header">Content Preview</div>', unsafe_allow_html=True)
    
    # Create tabs for each channel
    if st.session_state.funnel_content:
        tabs = st.tabs(list(st.session_state.funnel_content.keys()))
        
        for i, channel in enumerate(st.session_state.funnel_content.keys()):
            with tabs[i]:
                # Display stages as expandable sections
                for stage in ["Brand", "Demand Gen", "Demand Capture"]:
                    if stage in st.session_state.funnel_content[channel]:
                        with st.expander(f"{stage} Content", expanded=(stage == "Brand")):
                            # Display posts
                            posts = st.session_state.funnel_content[channel][stage]
                            if posts and isinstance(posts, list):
                                for j, post in enumerate(posts):
                                    if not isinstance(post, dict):
                                        continue
                                        
                                    st.markdown(f"#### Post {j+1}")
                                    
                                    # Create columns for post details
                                    cols = st.columns([3, 2])
                                    
                                    # Left column: Post Copy
                                    with cols[0]:
                                        st.markdown("**Post Copy:**")
                                        st.markdown(f"<div style='background-color:#f8f9fa;padding:10px;border-radius:5px;'>{post.get('Post Copy', '')}</div>", unsafe_allow_html=True)
                                    
                                    # Right column: Other details
                                    with cols[1]:
                                        if "Value Proposition" in post and post["Value Proposition"]:
                                            st.markdown("**Value Proposition:**")
                                            st.markdown(f"<div style='background-color:#f8f9fa;padding:10px;border-radius:5px;'>{post.get('Value Proposition', '')}</div>", unsafe_allow_html=True)
                                        
                                        if "Image Copy" in post and post["Image Copy"]:
                                            st.markdown("**Image Copy:**")
                                            st.markdown(f"<div style='background-color:#f8f9fa;padding:10px;border-radius:5px;'>{post.get('Image Copy', '')}</div>", unsafe_allow_html=True)
                                        
                                        if "CTA" in post and post["CTA"]:
                                            st.markdown("**Call to Action:**")
                                            st.markdown(f"<div style='background-color:#f8f9fa;padding:10px;border-radius:5px;'>{post.get('CTA', '')}</div>", unsafe_allow_html=True)
                                        
                                        if "Link" in post and post["Link"]:
                                            st.markdown("**Link:**")
                                            st.markdown(f"<div style='background-color:#f8f9fa;padding:10px;border-radius:5px;'>{post.get('Link', '')}</div>", unsafe_allow_html=True)
                                        
                                        if "Type" in post and post["Type"]:
                                            st.markdown(f"**Type:** {post.get('Type', '')}")
                                    
                                    st.markdown("---")
                            else:
                                st.info(f"No {stage} posts found for {channel}")

# Display main content area with instructions or results
main_content = st.container()

with main_content:
    # Always show this section if no content has been generated yet
    if not st.session_state.content_generated:
        st.markdown('<div class="sub-header">How It Works</div>', unsafe_allow_html=True)
        
        st.markdown("""
        <div class="instruction-message">
            <h3>👈 Fill in the sidebar and click Generate to begin</h3>
            <p>Complete all the required fields in the sidebar to generate your marketing funnel content.</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Feature boxes to fill the empty space
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            <div class="feature-box">
                <h3>🎯 Complete Marketing Funnel</h3>
                <p>Generate content for all stages of your marketing funnel:</p>
                <ul>
                    <li><b>Brand</b>: Establish your client's unique value proposition</li>
                    <li><b>Demand Gen</b>: Educate and nurture potential leads</li>
                    <li><b>Demand Capture</b>: Convert leads with compelling CTAs</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("""
            <div class="feature-box">
                <h3>📱 Multi-Channel Support</h3>
                <p>Create tailored content for various social platforms:</p>
                <ul>
                    <li><b>LinkedIn</b>: Professional, business-focused content</li>
                    <li><b>Facebook</b>: Engaging content with appropriate emojis</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)
            
        with col2:
            st.markdown("""
            <div class="feature-box">
                <h3>🤖 AI-Powered Content</h3>
                <p>Leverage OpenAI's models to generate:</p>
                <ul>
                    <li><b>Post Copy</b>: Engaging text tailored to each platform</li>
                    <li><b>Image Copy</b>: Suggestions for accompanying visuals</li>
                    <li><b>CTAs</b>: Compelling calls-to-action</li>
                    <li><b>Value Propositions</b>: Clear statements of client value</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("""
            <div class="feature-box">
                <h3>📊 Ready-to-Use Format</h3>
                <p>Export your content to Excel with:</p>
                <ul>
                    <li>Organized sheets for each platform</li>
                    <li>Clearly structured funnel stages</li>
                    <li>Formatted for easy sharing with clients</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)

    # Main app logic - handle the generate button
    if generate_button:
        if not api_key:
            st.error("Please enter your OpenAI API key.")
        elif not website_url:
            st.error("Please enter the client's website URL.")
        elif not channel_options:
            st.error("Please select at least one social media channel.")
        else:
            with st.spinner("Generating content funnel..."):
                success = generate_content()
                if success:
                    st.success("Content funnel successfully generated!")

    # Display results if content has been generated
    if st.session_state.content_generated:
        # Success message
        st.markdown(f"""
        <div class="success-message">
            <h3>✅ Content Generation Complete!</h3>
            <p>Processing time: {int(st.session_state.processing_time // 60)} minutes and {int(st.session_state.processing_time % 60)} seconds</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Download button for Excel
        if st.session_state.excel_data:
            company_name = st.session_state.generator.website_data['company_name'] if st.session_state.generator.website_data else "company"
            filename = f"{company_name}_marketing_funnel.xlsx"
            
            st.markdown(f"""
            <div class="info-message">
                <h3>📊 Download Your Marketing Funnel</h3>
                <p>Your content has been formatted into an Excel spreadsheet ready for use.</p>
                {get_excel_download_link(st.session_state.excel_data, filename)}
            </div>
            """, unsafe_allow_html=True)
        
        # Display content preview
        display_content_preview()

# Footer
st.markdown("---")
st.markdown("Made with ❤️ by Lead Engine Content Generator | Powered by OpenAI")
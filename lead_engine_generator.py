import openai
import pandas as pd
import os
import json
import re
import requests
import PyPDF2
import io
from bs4 import BeautifulSoup
from urllib.parse import urlparse
import time
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

class LeadEngineGenerator:
    def __init__(self, api_key):
        self.client = openai.OpenAI(api_key=api_key)
        self.extracted_data = {}
        self.website_data = None
        self.pdf_content = None
        self.lead_purpose = None
        self.asset_link = None
        self.funnel_content = {}
        # Add caching flags
        self._website_scraped = False
        self._pdf_extracted = False
    
    def scrape_website(self, url, force_rescrape=False):
        """Scrape content from a website URL"""
        # Return cached data if available and not forcing rescrape
        if self._website_scraped and self.website_data and not force_rescrape:
            print(f"Using cached website data for {url}")
            return self.website_data
            
        try:
            # Add http:// if missing
            if not url.startswith(('http://', 'https://')):
                url = 'https://' + url
                
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            }
            
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()  # Raise exception for 4XX/5XX responses
            
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Extract company name from domain
            domain = urlparse(url).netloc
            company_name = domain.split('.')[-2].capitalize()
            
            # Extract key website information
            data = {
                'company_name': company_name,
                'domain': domain,
                'url': url,
                'title': soup.title.text.strip() if soup.title else "",
                'meta_description': "",
                'h1_tags': [],
                'main_content': ""
            }
            
            # Get meta description
            meta_desc = soup.find('meta', attrs={'name': 'description'})
            if meta_desc:
                data['meta_description'] = meta_desc.get('content', '')
                
            # Get H1 tags
            for h1 in soup.find_all('h1'):
                if h1.text.strip():
                    data['h1_tags'].append(h1.text.strip())
            
            # Try to get main content
            main_content = soup.find('main') or soup.find(id=re.compile('(content|main)', re.I))
            if main_content:
                data['main_content'] = main_content.get_text(" ", strip=True)[:3000]  # Increased limit
            else:
                # Fallback: Take all paragraph text from the body
                paragraphs = [p.get_text(" ", strip=True) for p in soup.find_all('p')]
                data['main_content'] = " ".join(paragraphs)[:3000]
            
            self.website_data = data
            self._website_scraped = True  # Set the flag
            return data
            
        except Exception as e:
            print(f"Error scraping website: {str(e)}")
            return None
    
    def extract_pdf_content(self, pdf_url, force_reextract=False):
        """Extract text content from a PDF URL"""
        # Return cached data if available and not forcing re-extraction
        if self._pdf_extracted and self.pdf_content and not force_reextract:
            print(f"Using cached PDF content for {pdf_url}")
            return self.pdf_content
            
        try:
            # Add http:// if missing
            if not pdf_url.startswith(('http://', 'https://')):
                pdf_url = 'https://' + pdf_url
                
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            }
            
            response = requests.get(pdf_url, headers=headers, timeout=30)
            response.raise_for_status()
            
            # Use PyPDF2 to extract text
            pdf_file = io.BytesIO(response.content)
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            
            text = ""
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                text += page.extract_text() + "\n"
            
            # Limit content to avoid token issues with API
            text = text[:5000]
            self.pdf_content = text
            self._pdf_extracted = True  # Set the flag
            
            print(f"Successfully extracted PDF content ({len(text)} characters)")
            return text
            
        except Exception as e:
            print(f"Error extracting PDF content: {str(e)}")
            return None
    
    def set_lead_purpose(self, purpose, asset_link=None):
        """Set the lead generation purpose and asset link"""
        valid_purposes = ["Demo Booking", "Sales Meeting"]
        if purpose not in valid_purposes:
            raise ValueError(f"Purpose must be one of: {', '.join(valid_purposes)}")
        
        self.lead_purpose = purpose
        self.asset_link = asset_link
    
    def generate_funnel_content(self, website_url, pdf_url=None, num_posts=3, channels=None):
        """Generate content for the entire marketing funnel"""
        import time
        start_time = time.time()
        
        # Default channels if not specified
        if channels is None:
            channels = ["LinkedIn", "Facebook"]
        
        # Scrape website content - only if not already cached
        print(f"Preparing website data: {website_url}")
        website_data = self.scrape_website(website_url)
        if not website_data:
            return False
        
        # Extract PDF content if provided - only if not already cached
        pdf_content = None
        if pdf_url:
            print(f"Preparing PDF content from: {pdf_url}")
            pdf_content = self.extract_pdf_content(pdf_url)
        
        company_name = website_data['company_name']
        print(f"Generating content funnel for {company_name}...")
        
        # Store company data
        self.extracted_data[company_name] = {}
        
        # Initialize funnel content dictionary for each channel
        for channel in channels:
            self.funnel_content[channel] = {}
        
        # Generate content for each funnel stage and channel
        funnel_stages = ["Brand", "Demand Gen", "Demand Capture"]
        
        for channel in channels:
            print(f"\nGenerating content for {channel}...")
            
            for stage in funnel_stages:
                print(f"  Generating {stage} content...")
                # Create prompt based on the funnel stage and channel
                prompt = self._create_funnel_stage_prompt(stage, num_posts, website_data, pdf_content, channel)
                
                # Get response from GPT
                try:
                    raw_response = self.chat_with_gpt(prompt)
                    data = self.extract_json(raw_response)
                    
                    # Convert to list if it's a dictionary
                    posts = data if isinstance(data, list) else [data]
                    
                    # Store in our data dictionary
                    if channel not in self.funnel_content:
                        self.funnel_content[channel] = {}
                    self.funnel_content[channel][stage] = posts
                    
                    print(f"    ✓ Generated {len(posts)} {stage} posts")
                    
                except Exception as e:
                    print(f"    ✗ Error generating {stage} content: {str(e)}")
                    print(f"    Raw response preview: {raw_response[:100]}...")
        
        # Calculate and display the elapsed time
        elapsed_time = time.time() - start_time
        minutes, seconds = divmod(elapsed_time, 60)
        print(f"\nContent generation completed in {int(minutes)} minutes and {int(seconds)} seconds")
        
        return True
    
    def _create_funnel_stage_prompt(self, stage, count, website_data, pdf_content=None, channel="LinkedIn"):
        """Create a tailored prompt for each funnel stage and channel"""
        company = website_data['company_name']
        website_url = website_data['url']
        
        # Base company context
        company_context = f"Company: {company}\n"
        company_context += f"Website: {website_url}\n"
        company_context += f"Website title: {website_data['title']}\n"
        company_context += f"Description: {website_data['meta_description']}\n"
        
        if website_data['h1_tags']:
            company_context += f"Key headings: {', '.join(website_data['h1_tags'])}\n"
        
        if website_data['main_content']:
            company_context += f"Website content: {website_data['main_content'][:500]}...\n"
        
        # Add PDF content if available
        pdf_context = ""
        if pdf_content:
            pdf_context = f"\nDownloadable PDF content: {pdf_content[:1000]}...\n"
        
        # Lead purpose context
        purpose_context = ""
        if self.lead_purpose:
            purpose_context = f"\nLead generation purpose: {self.lead_purpose}"
            if self.asset_link:
                purpose_context += f"\nAsset link: {self.asset_link}"
        
        # Channel-specific instructions
        channel_instructions = {
            "LinkedIn": "Keep content professional and business-focused. Use a formal but engaging tone appropriate for LinkedIn.",
            "Facebook": "Keep content professional and business-focused. Use a formal but engaging tone appropriate for LinkedIn. Include up to 5 emojis"
        }
        
        channel_instruction = channel_instructions.get(channel, "")
        
        # Different prompts for each funnel stage
        prompts = {
            "Brand": f"""Based on this information about {company}:
{company_context}{pdf_context}{purpose_context}

Create {count} {channel} posts for the BRAND AWARENESS stage of a marketing funnel. These posts should establish the company's value proposition and brand identity. Each post should include:

1. "Value Proposition": A concise statement of the unique value the company offers
2. "Post Copy": Engaging {channel} post text (300-500 characters) that establishes brand identity
3. "Image Copy": A brief caption that could appear on an accompanying image (up to 100 characters)

{channel_instruction}

The content should resonate with the company's industry and target audience, focusing on building brand awareness rather than directly selling products/services. Format the output as a JSON array with these fields.
""",

            "Demand Gen": f"""Based on this information about {company}:
{company_context}{pdf_context}{purpose_context}

Create {count*2} {channel} posts for the DEMAND GENERATION stage of a marketing funnel. These posts should be split into two categories:

A. DOWNLOAD ASSET posts, each including:
1. "Post Copy": Engaging {channel} post text (300-500 characters) promoting the downloadable resource
2. "Image Copy": A brief caption that could appear on an accompanying image (up to 100 characters)
3. "Link": URL for downloading the asset (use "{self.asset_link if self.asset_link else "[Asset Link Placeholder]"}")

B. NURTURE FLOW posts, each including:
1. "Post Copy": Engaging {channel} post text (300-500 characters) that educates and nurtures potential leads
2. "Image Copy": A brief caption that could appear on an accompanying image (up to 100 characters)
3. "CTA": A clear call-to-action (up to 100 characters)

{channel_instruction}

Create {count} posts for each category (total of {count*2} posts). The content should build on brand awareness and move prospects closer to considering the company's solutions. Format the output as a JSON array with these fields and a "Type" field indicating either "Download Asset" or "Nurture Flow".
""",

            "Demand Capture": f"""Based on this information about {company}:
{company_context}{pdf_context}{purpose_context}

Create {count} {channel} posts for the DEMAND CAPTURE stage of a marketing funnel. These posts should focus on converting leads for {self.lead_purpose if self.lead_purpose else "lead generation"}. Each post should include:

1. "Post Copy": Engaging {channel} post text (300-500 characters) with a strong conversion focus
2. "Image Copy": A brief caption that could appear on an accompanying image (up to 100 characters)
3. "CTA": A compelling call-to-action (up to 100 characters) specific to {self.lead_purpose if self.lead_purpose else "lead conversion"}
4. "Link": URL for the conversion action (use "{self.asset_link if self.asset_link else "[Conversion Link Placeholder]"}")

{channel_instruction}

The content should create urgency and clearly communicate the next steps for interested prospects. Format the output as a JSON array with these fields.
"""
        }
        
        return prompts.get(stage, f"Create {count} {channel} posts for {company} formatted as JSON.")
    
    def chat_with_gpt(self, prompt):
        """Send prompt to OpenAI API and get response"""
        try:
            resp = self.client.chat.completions.create(
                model="gpt-4.1-nano",  # Using cheaper model for testing
                messages=[{"role": "user", "content": prompt}],
                temperature=0.7,
                max_tokens=4000
            )
            return resp.choices[0].message.content
        except Exception as e:
            print(f"Error calling OpenAI API: {str(e)}")
            raise

    def extract_json(self, text):
        """Extract JSON from the model response"""
        # Try to find JSON using regex pattern
        json_pattern = r'```(?:json)?\s*(\[.*\]|\{.*\})\s*```'
        match = re.search(json_pattern, text, re.DOTALL)

        if match:
            # Extract JSON from code blocks
            json_str = match.group(1)
        else:
            # Try to extract JSON without code blocks
            # Find the first opening bracket ([ or {)
            start_idx = min((text.find('{'), text.find('[')), key=lambda x: float('inf') if x == -1 else x)
            if start_idx == -1:
                raise ValueError("No JSON found in the response")
                
            # Find the matching closing bracket
            if text[start_idx] == '{':
                end_char = '}'
            else:
                end_char = ']'
                
            # Find the last closing bracket
            end_idx = text.rfind(end_char) + 1
            if end_idx == 0:
                raise ValueError(f"No matching closing bracket {end_char} found")
                
            json_str = text[start_idx:end_idx]
        
        # Parse the JSON
        try:
            return json.loads(json_str)
        except json.JSONDecodeError:
            # Try to fix common JSON issues
            json_str = json_str.replace("'", '"')
            json_str = re.sub(r',\s*}', '}', json_str)
            json_str = re.sub(r',\s*]', ']', json_str)
            return json.loads(json_str)
    
    def export_to_excel(self):
        """Save generated content to Excel file with structured funnel format"""
        if not self.funnel_content:
            print("❌ No funnel content collected. Excel file not created.")
            return None
        
        # Create pandas ExcelWriter
        desktop_path = os.path.expanduser("~/Desktop")
        company_name = self.website_data['company_name'] if self.website_data else "company"
        output_file = os.path.join(desktop_path, f"{company_name}.xlsx")
        
        # Create Excel writer
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Process each channel
            for channel in self.funnel_content.keys():
                # Create DataFrame for channel sheet - remove Funnel Stage column
                channel_data = []
                
                # Add Brand posts
                if "Brand" in self.funnel_content[channel]:
                    # Add section header row
                    channel_data.append({
                        "Post #": "",
                        "Value Proposition": "",
                        "Post Copy": "",
                        "Image Copy": "",
                        "CTA": "",
                        "Link": "",
                        "Type": "Brand"  # Moved from Funnel Stage to Type
                    })
                    
                    # Better error handling for posts
                    brand_posts = self.funnel_content[channel]["Brand"]
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
                if "Demand Gen" in self.funnel_content[channel]:
                    # Add section header row
                    channel_data.append({
                        "Post #": "",
                        "Value Proposition": "",
                        "Post Copy": "",
                        "Image Copy": "",
                        "CTA": "",
                        "Link": "",
                        "Type": "Demand Gen"  # Moved from Funnel Stage to Type
                    })
                    
                    demand_gen_posts = self.funnel_content[channel]["Demand Gen"]
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
                if "Demand Capture" in self.funnel_content[channel]:
                    # Add section header row
                    channel_data.append({
                        "Post #": "",
                        "Value Proposition": "",
                        "Post Copy": "",
                        "Image Copy": "",
                        "CTA": "",
                        "Link": "",
                        "Type": "Demand Capture"  # Moved from Funnel Stage to Type
                    })
                    
                    demand_capture_posts = self.funnel_content[channel]["Demand Capture"]
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
                                "Type": self.lead_purpose if self.lead_purpose else "Demand Capture"
                            }
                            channel_data.append(row)
                
                # Create DataFrame and export to Excel
                df = pd.DataFrame(channel_data)
                df.to_excel(writer, sheet_name=channel, index=False)
                
                # Apply styling
                self._apply_excel_styling(writer.sheets[channel], df)
        
        print(f"✅ Excel saved to {output_file}")
        return output_file
    
    def _apply_excel_styling(self, worksheet, df):
        """Apply formatting to Excel worksheet"""
        # Define styles
        header_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")  # Black background
        header_font = Font(color="FFFFFF", bold=True)  # White text
        post_num_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")  # Black background
        post_num_font = Font(color="FFFFFF", bold=True)  # White text
        funnel_stage_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")  # Blue background
        funnel_stage_font = Font(color="FFFFFF", bold=True)  # White text
        border = Border(
            left=Side(border_style="thin"), 
            right=Side(border_style="thin"),
            top=Side(border_style="thin"),
            bottom=Side(border_style="thin")
        )
        
        # Apply header styling
        for col_num, column_title in enumerate(df.columns, 1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border
            
            # Adjust column width
            column_letter = get_column_letter(col_num)
            if column_title in ["Post Copy", "Value Proposition"]:
                worksheet.column_dimensions[column_letter].width = 40
            elif column_title in ["Image Copy", "CTA", "Link"]:
                worksheet.column_dimensions[column_letter].width = 25
            else:
                worksheet.column_dimensions[column_letter].width = 15
        
        # Apply funnel stage styling and borders to all cells
        for row_num in range(2, len(df) + 2):  # +2 because Excel is 1-indexed and we have a header row
            for col_num in range(1, len(df.columns) + 1):
                cell = worksheet.cell(row=row_num, column=col_num)
                cell.border = border
                
                # Apply section header styling (now using Type column for section headers)
                if col_num == len(df.columns) and cell.value in ["Brand", "Demand Gen", "Demand Capture"]:  # Type column
                    # Find the first cell in this row
                    first_cell = worksheet.cell(row=row_num, column=1)
                    first_cell.fill = funnel_stage_fill
                    first_cell.font = funnel_stage_font
                    first_cell.alignment = Alignment(horizontal='center', vertical='center')
                    # Merge cells for the section header
                    worksheet.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=len(df.columns))
                    # Set the value in the merged cell
                    first_cell.value = cell.value
                    cell.value = ""  # Clear the original cell
                
                # Apply styling to Post # column
                if col_num == 1 and cell.value and isinstance(cell.value, (int, float)):  # Post number column
                    cell.fill = post_num_fill
                    cell.font = post_num_font
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                
                # Set text wrapping for content cells
                content_columns = [col for col, title in enumerate(df.columns, 1) 
                                  if title in ["Value Proposition", "Post Copy", "Image Copy", "CTA"]]
                if col_num in content_columns:
                    cell.alignment = Alignment(wrapText=True, vertical='top')
        
        # Auto-fit row heights
        for row_num in range(2, len(df) + 2):
            max_text_lines = 1
            content_columns = [col for col, title in enumerate(df.columns, 1) 
                              if title in ["Value Proposition", "Post Copy", "Image Copy", "CTA"]]
            for col_num in content_columns:  # Check text length in main content columns
                cell_value = worksheet.cell(row=row_num, column=col_num).value
                if cell_value:
                    lines = len(str(cell_value).split('\n'))
                    max_text_lines = max(max_text_lines, lines)
                    # Also consider approximate wrapping
                    approx_lines = len(str(cell_value)) // 40 + 1  # 40 chars per line approx
                    max_text_lines = max(max_text_lines, approx_lines)
            
            # Set row height (15 points per line of text, minimum 20)
            row_height = max(20, 15 * max_text_lines)
            worksheet.row_dimensions[row_num].height = row_height
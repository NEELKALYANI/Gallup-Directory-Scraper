import requests
from bs4 import BeautifulSoup
import pandas as pd
import sys
import io
from datetime import datetime
from lxml import html
import time
import os

def scrape_gallup_profile(url, debug=False):
    # Send a request to the URL
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        print(f"Successfully fetched the URL with status code: {response.status_code}")
    except requests.exceptions.RequestException as e:
        print(f"Error fetching URL: {e}")
        return None
    
    # Parse the HTML content
    soup = BeautifulSoup(response.text, 'html.parser')
    
    # For XPath extraction
    tree = html.fromstring(response.content)
    
    # Initialize dictionary to store profile data with empty strings
    profile_data = {
        'URL': url,
        'Name': "",
        'Email': "N/A",  # Added email field
        'Country': "",
        'Expertise': "",
        'Availability': "N/A",
        'Method': "N/A",
        'Language': "N/A",
        'About Me': "N/A"
    }
    
    # Extract name
    name_element = soup.find('h1', class_='c-person__name')
    if name_element:
        profile_data['Name'] = name_element.text.strip()
        print(f"Found name: {profile_data['Name']}")
    else:
        print("Name element not found with class 'c-person__name'")
        # Try alternate selectors
        all_h1s = soup.find_all('h1')
        for h1 in all_h1s:
            if h1.text.strip():
                profile_data['Name'] = h1.text.strip()
                print(f"Using alternative name: {profile_data['Name']}")
                break
    
    # Extract email address
    email_element = soup.find('a', class_='c-person__email')
    if email_element:
        profile_data['Email'] = email_element.text.strip()
        print(f"Found email: {profile_data['Email']}")
    else:
        # Try alternative method to find email (look for mailto links)
        email_links = soup.find_all('a', href=lambda href: href and href.startswith('mailto:'))
        if email_links:
            profile_data['Email'] = email_links[0].text.strip()
            print(f"Found email from mailto link: {profile_data['Email']}")
    
    # Extract country
    if name_element:
        country_div = name_element.find_next('div')
        if country_div:
            country_text = country_div.text.strip()
            profile_data['Country'] = country_text
            print(f"Found country: {profile_data['Country']}")
    
    # Extract expertise using multiple approaches
    expertise_extracted = False
    
    # APPROACH 1: Try the original XPath (keep for backward compatibility)
    expertise_elements = tree.xpath('/html/body/div[2]/div/main/article/div[1]/header/div[2]/div[3]/p[1]/text()')
    if expertise_elements:
        profile_data['Expertise'] = expertise_elements[0].strip()
        print(f"Found expertise using original XPath: {profile_data['Expertise']}")
        expertise_extracted = True
    else:
        print("Expertise not found using original XPath")
        
        # APPROACH 2: Try more robust XPath patterns
        # Looking for elements with likely expertise text
        expertise_patterns = [
            # Try to find paragraphs containing expertise/specialty keywords
            "//p[contains(text(), 'Expertise')]",
            "//p[contains(text(), 'Specialty')]",
            "//p[contains(text(), 'Focus')]",
            "//div[contains(@class, 'expertise')]//p",
            "//div[contains(@class, 'specialty')]//p",
            # Try common class names or structures
            "//div[contains(@class, 'person-info')]//p",
            # Try to find elements near the name
            "//h1[contains(@class, 'person')]/following-sibling::div//p"
        ]
        
        for pattern in expertise_patterns:
            try:
                elements = tree.xpath(pattern)
                if elements:
                    potential_expertise = elements[0].text_content().strip()
                    # Filter out if it's just a label
                    if len(potential_expertise.split()) > 1 and not potential_expertise.startswith("Email"):
                        profile_data['Expertise'] = potential_expertise
                        print(f"Found expertise using pattern '{pattern}': {profile_data['Expertise']}")
                        expertise_extracted = True
                        break
            except Exception as e:
                print(f"Error with pattern '{pattern}': {e}")
    
    # APPROACH 3: Look for specific keywords in paragraphs (if still not found)
    if not expertise_extracted:
        print("Trying keyword-based expertise extraction...")
        expertise_keywords = ['expertise', 'specialty', 'specialist', 'specializes', 'specializing', 
                             'focus', 'focuses', 'focusing', 'experienced in', 'skilled in']
        
        # Look for paragraphs that might contain expertise info
        paragraphs = soup.find_all('p')
        for p in paragraphs[:5]:  # Check the first few paragraphs
            p_text = p.text.lower().strip()
            if any(keyword in p_text for keyword in expertise_keywords):
                profile_data['Expertise'] = p.text.strip()
                print(f"Found expertise using keyword search: {profile_data['Expertise']}")
                expertise_extracted = True
                break
    
    # APPROACH 4: Try to find expertise by examining page structure and field patterns
    if not expertise_extracted:
        # Look for strong elements that might be field labels
        strong_tags = soup.find_all(['strong', 'b', 'h2', 'h3', 'h4', 'h5', 'label'])
        for tag in strong_tags:
            tag_text = tag.text.strip().lower()
            if 'expertise' in tag_text or 'specialty' in tag_text or 'specialization' in tag_text:
                print(f"Found expertise label: '{tag.text.strip()}'")
                
                # Try different approaches to get the content
                # Similar to what you're already doing for other fields
                parent = tag.parent
                if parent:
                    content = parent.text.replace(tag.text, '').strip()
                    if content:
                        profile_data['Expertise'] = content
                        print(f"Extracted expertise from parent: '{content}'")
                        expertise_extracted = True
                        break
                
                # Look for next sibling or next element
                next_elem = tag.next_sibling
                while next_elem and isinstance(next_elem, str) and not next_elem.strip():
                    next_elem = next_elem.next_sibling
                
                if next_elem:
                    if hasattr(next_elem, 'text'):
                        content = next_elem.text.strip()
                    else:
                        content = next_elem.strip()
                    
                    if content:
                        profile_data['Expertise'] = content
                        print(f"Extracted expertise from next element: '{content}'")
                        expertise_extracted = True
                        break

    # If still not found and debug mode is on, save page structure for debugging
    if not expertise_extracted:
        print("WARNING: Could not extract Expertise using any method")
        if debug:
            # Create debug directory if it doesn't exist
            os.makedirs("debug", exist_ok=True)
            # Save HTML structure for debugging
            safe_name = ''.join(c if c.isalnum() else '_' for c in profile_data['Name']) if profile_data['Name'] else 'unknown'
            debug_file = f"debug/debug_{safe_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
            with open(debug_file, "w", encoding="utf-8") as f:
                f.write(response.text)
            print(f"Saved HTML structure to {debug_file} for inspection")
    
    # APPROACH 1: Direct search for specific field labels
    print("\n--- APPROACH 1: Searching for field labels ---")
    field_patterns = {
        'Availability': ['availability', 'available', 'schedule'],
        'Method': ['method', 'delivery', 'coaching method', 'coaching style'],
        'Language': ['language', 'languages', 'speaks']
    }
    
    # Look for strong elements that might be field labels
    strong_tags = soup.find_all(['strong', 'b', 'h2', 'h3', 'h4', 'h5', 'label'])
    print(f"Found {len(strong_tags)} potential field label elements")
    
    for tag in strong_tags:
        tag_text = tag.text.strip().lower()
        
        # Check if this tag matches any of our field patterns
        for field, patterns in field_patterns.items():
            if any(pattern in tag_text for pattern in patterns):
                print(f"Found {field} label: '{tag.text.strip()}'")
                
                # Try different approaches to get the content
                # 1. Check if content is in parent element after the strong tag
                parent = tag.parent
                if parent:
                    # Get text excluding the strong tag itself
                    content = parent.text.replace(tag.text, '').strip()
                    if content:
                        profile_data[field] = content
                        print(f"Extracted {field} from parent: '{content}'")
                        continue
                
                # 2. Look for next sibling or next element
                next_elem = tag.next_sibling
                # Skip empty text nodes
                while next_elem and isinstance(next_elem, str) and not next_elem.strip():
                    next_elem = next_elem.next_sibling
                
                if next_elem:
                    if hasattr(next_elem, 'text'):
                        content = next_elem.text.strip()
                    else:
                        content = next_elem.strip()
                    
                    if content:
                        profile_data[field] = content
                        print(f"Extracted {field} from next element: '{content}'")
                        continue
                
                # 3. Look for paragraph after the containing element
                next_p = tag.find_next('p')
                if next_p:
                    content = next_p.text.strip()
                    if content:
                        profile_data[field] = content
                        print(f"Extracted {field} from next paragraph: '{content}'")
                        continue
    
    # Extract About Me section
    about_section = soup.find('div', class_='c-person--content')
    if about_section:
        profile_data['About Me'] = about_section.text.strip()
        print(f"Found About Me section with length: {len(profile_data['About Me'])}")
    
    # Check if any fields are still missing
    for field, value in profile_data.items():
        if value == "N/A" and field != 'URL':
            print(f"Warning: Could not extract {field}")
    
    return profile_data

def process_urls_from_excel(input_file='links.xlsx', url_column='URL', debug=False):
    """
    Read URLs from an Excel file and process each one
    """
    try:
        # Read the Excel file
        print(f"Reading URLs from {input_file}...")
        df = pd.read_excel(input_file)
        
        if url_column not in df.columns:
            print(f"Error: Column '{url_column}' not found in {input_file}")
            print(f"Available columns: {', '.join(df.columns)}")
            return []
        
        # Filter out invalid URLs
        urls = df[url_column].dropna().tolist()
        valid_urls = [url for url in urls if isinstance(url, str) and url.startswith('http')]
        
        print(f"Found {len(valid_urls)} valid URLs to process")
        
        # Process each URL
        all_profiles = []
        for i, url in enumerate(valid_urls):
            print(f"\n[{i+1}/{len(valid_urls)}] Processing: {url}")
            profile_data = scrape_gallup_profile(url, debug=debug)
            if profile_data:
                all_profiles.append(profile_data)
            
            # Add a delay between requests to be respectful to the server
            if i < len(valid_urls) - 1:
                delay = 4 + (i % 3)  # Vary the delay slightly to appear more human-like
                print(f"Waiting {delay} seconds before next request...")
                time.sleep(delay)
        
        return all_profiles
        
    except Exception as e:
        print(f"Error processing URLs from Excel: {e}")
        return []

def save_to_excel(profile_data_list, output_file=None):

    if not profile_data_list:
        print("No data to save")
        return
    
    # Create DataFrame and save to Excel
    df = pd.DataFrame(profile_data_list)
    try:
        df.to_excel(output_file, index=False)
        print(f"\nProfile data saved to {output_file}")
        
        # Print a summary
        print(f"\nSuccessfully scraped {len(profile_data_list)} profiles")
        
        # Print field completion statistics
        for field in df.columns:
            if field != 'URL':
                missing = df[df[field] == 'N/A'].shape[0]
                pct_complete = (len(profile_data_list) - missing) / len(profile_data_list) * 100
                print(f"{field}: {pct_complete:.1f}% complete ({missing} missing)")
            
    except Exception as e:
        print(f"Error saving to Excel: {e}")

def main():

    print("Gallup Profile Scraper")
    print("=====================")
    
    # Get input file from user or use default
    input_file ="links.xlsx"
    
    # Get URL column name
    url_column = "URL"
    
    # Get output file name
    output_file = "Scraped2.xlsx"
    
    # Process URLs
    profile_data_list = process_urls_from_excel(input_file, url_column)
    
    if profile_data_list:
        save_to_excel(profile_data_list, output_file)
    else:
        print("Failed to extract any profile data")

if __name__ == "__main__":
    main()
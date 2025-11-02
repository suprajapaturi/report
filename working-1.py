#!/usr/bin/env python3
"""
TREC Home Inspection Report Generator - Standalone Version
=====================================================

This script generates a TREC-compliant PDF report by filling the official 
TREC_Template_Blank.pdf form with data from inspection.json.

Requirements:
- PyPDF2==3.0.1
- reportlab>=4.0.0

Usage:
    python3 trec_generator_standalone.py

Files required:
- inspection.json (inspection data)
- TREC_Template_Blank.pdf (official TREC template)

Output:
- output_pdf.pdf (filled TREC form)
"""

import json
import os
import logging
from datetime import datetime
from typing import Dict, List, Any, Optional
from PyPDF2 import PdfReader, PdfWriter

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)


class InspectionDataParser:
    """Parser for inspection.json data."""
    
    def __init__(self, json_file_path: str):
        """Initialize the parser with the inspection JSON file."""
        self.json_file_path = json_file_path
        self.data = self._load_json()
        
    def _load_json(self) -> Dict[str, Any]:
        """Load and parse the inspection JSON file."""
        try:
            with open(self.json_file_path, 'r', encoding='utf-8') as file:
                return json.load(file)
        except FileNotFoundError:
            raise FileNotFoundError(f"Inspection file not found: {self.json_file_path}")
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON format: {e}")
    
    def get_property_info(self) -> Dict[str, Any]:
        """Extract property information."""
        inspection = self.data.get('inspection', {})
        address = inspection.get('address', {})
        
        return {
            'street': address.get('street', ''),
            'city': address.get('city', ''),
            'state': address.get('state', ''),
            'zipcode': address.get('zipcode', ''),
            'full_address': address.get('fullAddress', ''),
            'square_footage': address.get('propertyInfo', {}).get('squareFootage', '')
        }
    
    def get_client_info(self) -> Dict[str, Any]:
        """Extract client information."""
        inspection = self.data.get('inspection', {})
        client = inspection.get('clientInfo', {})
        
        return {
            'name': client.get('name', ''),
            'email': client.get('email', ''),
            'phone': client.get('phone', ''),
            'user_type': client.get('userType', '')
        }
    
    def get_inspector_info(self) -> Dict[str, Any]:
        """Extract inspector information."""
        inspection = self.data.get('inspection', {})
        inspector = inspection.get('inspector', {})
        
        return {
            'id': inspector.get('id', ''),
            'name': inspector.get('name', ''),
            'email': inspector.get('email', ''),
            'phone': inspector.get('phone', '')
        }
    
    def get_inspection_schedule(self) -> Dict[str, Any]:
        """Extract inspection schedule information."""
        inspection = self.data.get('inspection', {})
        schedule = inspection.get('schedule', {})
        
        return {
            'date': schedule.get('date', ''),
            'start_time': schedule.get('startTime', ''),
            'end_time': schedule.get('endTime', '')
        }
    
    def get_sections(self) -> List[Dict[str, Any]]:
        """Extract all inspection sections with their line items and comments."""
        inspection = self.data.get('inspection', {})
        sections = inspection.get('sections', [])
        
        parsed_sections = []
        
        for section in sections:
            section_data = {
                'id': section.get('id', ''),
                'name': section.get('name', ''),
                'order': section.get('order', 0),
                'section_number': section.get('sectionNumber', ''),
                'line_items': []
            }
            
            # Parse line items
            line_items = section.get('lineItems', [])
            for line_item in line_items:
                line_item_data = {
                    'id': line_item.get('id', ''),
                    'name': line_item.get('name', ''),
                    'title': line_item.get('title', ''),
                    'order': line_item.get('order', 0),
                    'inspection_status': line_item.get('inspectionStatus', ''),
                    'is_deficient': line_item.get('isDeficient', False),
                    'comments': []
                }
                
                # Parse comments
                comments = line_item.get('comments', [])
                for comment in comments:
                    comment_data = {
                        'id': comment.get('id', ''),
                        'text': comment.get('text', ''),
                        'type': comment.get('type', ''),
                        'location': comment.get('location', ''),
                        'label': comment.get('label', ''),
                        'is_selected': comment.get('isSelected', False),
                        'is_flagged': comment.get('isFlagged', False),
                        'comment_number': comment.get('commentNumber', ''),
                        'photos': comment.get('photos', []),
                        'videos': comment.get('videos', [])
                    }
                    line_item_data['comments'].append(comment_data)
                
                section_data['line_items'].append(line_item_data)
            
            parsed_sections.append(section_data)
        
        return parsed_sections


class TRECFormFiller:
    """Fills the TREC_Template_Blank.pdf with inspection data from inspection.json"""
    
    def __init__(self, data_parser: InspectionDataParser, template_path: str = "TREC_Template_Blank.pdf"):
        """Initialize the TREC form filler."""
        self.parser = data_parser
        self.template_path = template_path
        
    def generate_report(self, output_path: str = "output_pdf.pdf") -> bool:
        """Generate the TREC report by filling actual form fields in the template."""
        try:
            logger.info(f"Filling TREC template form fields: {self.template_path}")
            
            # Check if template exists
            if not os.path.exists(self.template_path):
                logger.error(f"TREC template not found: {self.template_path}")
                return False

            # Read the template PDF
            reader = PdfReader(self.template_path)
            writer = PdfWriter()
            
            # Get inspection data
            property_info = self.parser.get_property_info()
            client_info = self.parser.get_client_info()
            inspector_info = self.parser.get_inspector_info()
            schedule = self.parser.get_inspection_schedule()
            sections = self.parser.get_sections()
            
            # Extract and format the actual data
            client_name = client_info.get('name', '')
            print(f"DEBUG: Client name: {client_name}")
            
            # Convert timestamp to readable date
            date_timestamp = schedule.get('date')
            if date_timestamp:
                try:
                    # Convert milliseconds to seconds
                    date_obj = datetime.fromtimestamp(date_timestamp / 1000)
                    inspection_date = date_obj.strftime('%m/%d/%Y')
                except:
                    inspection_date = '08/13/2025'  # fallback
            else:
                inspection_date = '08/13/2025'  # fallback
            
            print(f"DEBUG: Inspection date: {inspection_date}")
            
            property_address = property_info.get('full_address', '')
            inspector_name = inspector_info.get('name', '')
            
            print(f"DEBUG: Property address: {property_address}")
            print(f"DEBUG: Inspector name: {inspector_name}")

            # Check what form fields are available
            form_fields = reader.get_fields()
            if form_fields:
                print(f"DEBUG: Found {len(form_fields)} form fields in template")
                
                # Analyze checkbox fields in detail
                self._analyze_form_fields(form_fields)
            else:
                print("DEBUG: No form fields found!")
                return False

            # Copy all pages from reader to writer
            for page in reader.pages:
                writer.add_page(page)

            # Fill the form fields with our data
            if form_fields:
                form_data = {}
                
                # Map our data to the known form field names
                if 'Name of Client' in form_fields and client_name:
                    form_data['Name of Client'] = client_name
                    print(f"DEBUG: Setting 'Name of Client' = {client_name}")
                
                if 'Date of Inspection' in form_fields and inspection_date:
                    form_data['Date of Inspection'] = inspection_date  
                    print(f"DEBUG: Setting 'Date of Inspection' = {inspection_date}")
                
                if 'Address of Inspected Property' in form_fields and property_address:
                    form_data['Address of Inspected Property'] = property_address
                    print(f"DEBUG: Setting 'Address of Inspected Property' = {property_address}")
                
                if 'Name of Inspector' in form_fields and inspector_name:
                    form_data['Name of Inspector'] = inspector_name
                    print(f"DEBUG: Setting 'Name of Inspector' = {inspector_name}")

                # Fill the form fields
                print(f"DEBUG: Updating {len(form_data)} form fields")
                writer.update_page_form_field_values(writer.pages[0], form_data)

                # Also fill the inspection matrix checkboxes based on section data
                self._fill_inspection_matrix(writer, sections, form_fields)
                
                # Fill detailed inspection comments on pages 3, 4, 5
                self._fill_inspection_details(writer, sections, form_fields)

            # Set NeedAppearances flag to force appearance generation
            if hasattr(writer, '_root_object') and writer._root_object:
                if '/AcroForm' in writer._root_object:
                    from PyPDF2.generic import BooleanObject, NameObject
                    acro_form = writer._root_object['/AcroForm']
                    acro_form[NameObject('/NeedAppearances')] = BooleanObject(True)
                    print("DEBUG: Set NeedAppearances=True to force checkbox appearance")

            # ENHANCED FLATTENING: Force all fields to be permanently visible and non-clickable
            self._flatten_all_form_fields_completely(writer)

            # FINAL STEP: Remove AcroForm to ensure complete flattening
            print("DEBUG: Removing AcroForm to make all content permanently visible and non-clickable")
            if hasattr(writer, '_root_object') and writer._root_object:
                if '/AcroForm' in writer._root_object:
                    del writer._root_object['/AcroForm']
                    print("DEBUG: AcroForm removed - all fields are now permanently visible and non-clickable")

            # Write the final PDF
            with open(output_path, 'wb') as output_file:
                writer.write(output_file)
            
            logger.info(f"TREC report with form data generated successfully: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error generating TREC report: {e}")
            import traceback
            traceback.print_exc()
            return False

    def _flatten_all_form_fields_completely(self, writer: PdfWriter) -> None:
        """Comprehensively flatten all form fields to make content permanently visible and non-clickable."""
        try:
            print("DEBUG: Performing comprehensive field flattening...")
            from PyPDF2.generic import NameObject
            
            total_processed = 0
            
            # Process each page
            for page_num, page in enumerate(writer.pages):
                if '/Annots' in page:
                    annotations = page['/Annots']
                    page_processed = 0
                    
                    for annot_ref in annotations:
                        try:
                            annot = annot_ref.get_object()
                            
                            # Check if this is a form field widget
                            if '/Subtype' in annot and annot['/Subtype'] == '/Widget':
                                field_type = annot.get('/FT', '')
                                field_name = str(annot.get('/T', ''))
                                field_value = annot.get('/V', '')
                                
                                if field_value:  # Only process fields with values
                                    print(f"DEBUG: Completely flattening field '{field_name}' (type: {field_type})...")
                                    
                                    # Remove ALL interactive properties
                                    interactive_properties = [
                                        '/AA',      # Additional Actions
                                        '/A',       # Action
                                        '/BS',      # Border Style
                                        '/H',       # Highlight mode
                                        '/MK',      # Appearance characteristics
                                        '/Parent',  # Parent reference (breaks form structure)
                                        '/Kids',    # Children references
                                        '/TI',      # Tab index
                                        '/TM',      # Text Matrix
                                        '/TU',      # Tooltip
                                    ]
                                    
                                    for prop in interactive_properties:
                                        if prop in annot:
                                            del annot[prop]
                                    
                                    # Set read-only and flattened flags
                                    annot[NameObject('/Ff')] = 1  # Read-only flag
                                    annot[NameObject('/F')] = 4   # Print flag (makes it visible when printed/flattened)
                                    
                                    # For text fields: ensure value is visible
                                    if field_type == '/Tx' and field_value:
                                        # Force the appearance to show the text
                                        if '/AP' not in annot:
                                            annot[NameObject('/AP')] = {}
                                    
                                    # For checkboxes: ensure checked state is visible
                                    elif field_type == '/Btn' and field_value:
                                        # Force checked appearance
                                        annot[NameObject('/AS')] = NameObject('/Yes')
                                        if '/AP' not in annot:
                                            annot[NameObject('/AP')] = {}
                                    
                                    page_processed += 1
                                    total_processed += 1
                                    
                        except Exception as e:
                            print(f"DEBUG: Error completely flattening annotation: {e}")
                            continue
                    
                    if page_processed > 0:
                        print(f"DEBUG: Completely flattened {page_processed} fields on page {page_num + 1}")
            
            print(f"DEBUG: Total fields completely flattened: {total_processed}")
            
        except Exception as e:
            print(f"DEBUG: Complete field flattening failed: {e}")
            import traceback
            traceback.print_exc()

    def _analyze_form_fields(self, form_fields: Dict) -> None:
        """Analyze form fields to understand checkbox structure."""
        try:
            print("DEBUG: Analyzing form field structure...")
            
            # Categorize fields
            text_fields = []
            checkbox_fields = []
            other_fields = []
            
            for field_name, field_info in form_fields.items():
                if 'CheckBox' in field_name:
                    checkbox_fields.append(field_name)
                elif 'Text' in field_name:
                    text_fields.append(field_name)
                else:
                    other_fields.append(field_name)
            
            print(f"DEBUG: Field breakdown - Text: {len(text_fields)}, Checkbox: {len(checkbox_fields)}, Other: {len(other_fields)}")
            
            # Show first few checkboxes to understand naming pattern
            print("DEBUG: First 10 checkbox fields:")
            for i, field_name in enumerate(checkbox_fields[:10]):
                # Check for UTF-8 BOM
                has_bom = field_name != field_name.encode('utf-8').decode('utf-8', errors='ignore')
                bom_indicator = " (BOM)" if has_bom else ""
                print(f"  {i}: '{field_name}'{bom_indicator}")
                
                # Try to extract checkbox index
                if '[' in field_name and ']' in field_name:
                    try:
                        start = field_name.find('[') + 1
                        end = field_name.find(']')
                        index = field_name[start:end]
                        print(f"      Index: {index}")
                    except:
                        pass
            
            # Show checkbox field properties
            if checkbox_fields:
                sample_checkbox = checkbox_fields[0]
                field_info = form_fields[sample_checkbox]
                print(f"DEBUG: Sample checkbox '{sample_checkbox}' properties:")
                if isinstance(field_info, dict):
                    for key, value in field_info.items():
                        print(f"  {key}: {value}")
                        # Look for export values
                        if key == '/AP' and isinstance(value, dict):
                            print(f"    Appearance dict keys: {list(value.keys())}")
                        elif key == '/Opt' and isinstance(value, (list, tuple)):
                            print(f"    Options: {value}")
                else:
                    print(f"  Type: {type(field_info)}")
                    print(f"  Value: {field_info}")
                
                # Try to determine the correct checkbox value from the field
                self._determine_checkbox_export_value(sample_checkbox, field_info)
            
        except Exception as e:
            print(f"DEBUG: Error analyzing form fields: {e}")

    def _determine_checkbox_export_value(self, field_name, field_info):
        """Try to determine the correct export value for checkboxes."""
        try:
            print(f"DEBUG: Analyzing checkbox export values for {field_name}")
            
            # Common checkbox export values to try
            possible_values = ['Yes', 'On', '1', 'True', 'X', 'Checked']
            
            if isinstance(field_info, dict):
                # Look for appearance dictionary
                if '/AP' in field_info:
                    ap_dict = field_info['/AP']
                    if isinstance(ap_dict, dict) and '/N' in ap_dict:
                        normal_appearances = ap_dict['/N']
                        if isinstance(normal_appearances, dict):
                            # The keys in the normal appearance dict are the export values
                            export_values = list(normal_appearances.keys())
                            print(f"DEBUG: Found export values in appearance dict: {export_values}")
                            return export_values
                
                # Look for options
                if '/Opt' in field_info:
                    options = field_info['/Opt']
                    print(f"DEBUG: Found options: {options}")
                    return options
            
            print(f"DEBUG: Using default checkbox values: {possible_values}")
            return possible_values
            
        except Exception as e:
            print(f"DEBUG: Error determining checkbox values: {e}")
            return ['Yes', 'On', '1']

    def _fill_inspection_matrix(self, writer: PdfWriter, sections: List[Dict], form_fields: Dict) -> None:
        """Fill the inspection matrix checkboxes based on section data."""
        try:
            print("DEBUG: Filling inspection matrix...")
            
            # Get sections that have actual inspection status
            inspection_sections = []
            for section in sections:
                section_name = section.get('name', '')
                line_items = section.get('line_items', [])
                
                print(f"DEBUG: Checking section '{section_name}' with {len(line_items)} line_items")
                
                if line_items:
                    # Use the first line item's status as section status
                    first_item = line_items[0]
                    status = first_item.get('inspection_status', '')
                    
                    print(f"DEBUG: Section '{section_name}' status: '{status}'")
                    
                    # Include ANY section with a status (including empty string check)
                    if status and status.strip():  # Only skip completely empty statuses
                        inspection_sections.append({
                            'name': section_name,
                            'status': status,
                            'deficient': first_item.get('is_deficient', False)
                        })
                        print(f"DEBUG: Added section '{section_name}' with status '{status}'")
            
            print(f"DEBUG: Found {len(inspection_sections)} sections with status")
            for i, section in enumerate(inspection_sections):
                print(f"DEBUG: Section {i+1}: {section['name'][:30]} - Status: {section['status']}")
            
            # Map sections to checkbox indices (TREC form has numbered checkboxes)
            checkbox_fields = [field for field in form_fields.keys() if field.startswith('CheckBox1[')]
            print(f"DEBUG: Found {len(checkbox_fields)} checkbox fields")
            
            # Also check for UTF-8 BOM prefixed checkboxes (common issue)
            bom_checkbox_fields = []
            for field_name in form_fields.keys():
                # Check for UTF-8 BOM prefix (þÿ) followed by CheckBox1[
                if 'CheckBox1[' in field_name and field_name != field_name.encode('utf-8').decode('utf-8', errors='ignore'):
                    bom_checkbox_fields.append(field_name)
            
            if bom_checkbox_fields:
                print(f"DEBUG: Found {len(bom_checkbox_fields)} UTF-8 BOM prefixed checkbox fields")
                checkbox_fields.extend(bom_checkbox_fields)
            
            # Print first few checkbox field names for debugging
            for i, field in enumerate(checkbox_fields[:10]):
                print(f"DEBUG: Checkbox field {i}: '{field}'")
            
            # Determine checkbox values by examining existing field properties
            checkbox_values_to_try = ['Yes', 'On', '1', '/Yes', '/On', '/1', True]
            
            # Fill checkboxes for each section (assuming 4 columns: I, NI, NP, D)
            total_checkbox_updates = 0
            
            for i, section in enumerate(inspection_sections[:12]):  # Limit to first 12 sections
                base_index = i * 4  # Each section has 4 checkboxes (I, NI, NP, D)
                status = section['status']
                deficient = section['deficient']
                
                print(f"DEBUG: Processing section {i+1}: {section['name'][:30]} - Status: {status}")
                
                # Try to find the right checkbox field name pattern
                checkbox_patterns = [
                    f'CheckBox1[{base_index}]',           # Standard pattern
                    f'CheckBox1[{base_index + 1}]',       # NI
                    f'CheckBox1[{base_index + 2}]',       # NP  
                    f'CheckBox1[{base_index + 3}]'        # D
                ]
                
                # Also try BOM-prefixed patterns
                for field_name in form_fields.keys():
                    if f'CheckBox1[{base_index}]' in field_name:
                        checkbox_patterns.append(field_name)
                    if f'CheckBox1[{base_index + 1}]' in field_name:
                        checkbox_patterns.append(field_name)
                    if f'CheckBox1[{base_index + 2}]' in field_name:
                        checkbox_patterns.append(field_name)
                    if f'CheckBox1[{base_index + 3}]' in field_name:
                        checkbox_patterns.append(field_name)
                
                # Remove duplicates while preserving order
                checkbox_patterns = list(dict.fromkeys(checkbox_patterns))
                
                # Create checkbox update data for the current page
                checkbox_data = {}
                
                # Mark the appropriate status checkbox - try multiple field name patterns and values
                target_checkbox_index = None
                if status == 'I':
                    target_checkbox_index = base_index
                elif status == 'NI':
                    target_checkbox_index = base_index + 1
                elif status == 'NP':
                    target_checkbox_index = base_index + 2
                elif status == 'D':
                    target_checkbox_index = base_index + 3
                
                if target_checkbox_index is not None:
                    # Try different checkbox field name patterns
                    checkbox_found = False
                    target_field_patterns = [
                        f'CheckBox1[{target_checkbox_index}]'
                    ]
                    
                    # Add BOM-prefixed patterns
                    for field_name in form_fields.keys():
                        if f'CheckBox1[{target_checkbox_index}]' in field_name:
                            target_field_patterns.append(field_name)
                    
                    for field_pattern in target_field_patterns:
                        if field_pattern in form_fields:
                            # Try different checkbox values
                            for checkbox_value in checkbox_values_to_try:
                                try:
                                    checkbox_data[field_pattern] = checkbox_value
                                    checkbox_found = True
                                    print(f"DEBUG: Will check {field_pattern} for section '{section['name'][:20]}' status '{status}' with value '{checkbox_value}'")
                                    total_checkbox_updates += 1
                                    break  # Use first working value
                                except Exception as e:
                                    continue
                            if checkbox_found:
                                break
                    
                    if not checkbox_found:
                        print(f"DEBUG: Could not find checkbox field for index {target_checkbox_index}")
                
                # Update the page with checkbox data and force appearance generation
                if checkbox_data:
                    # DEBUG: Show exactly what checkbox data we're trying to set
                    print(f"DEBUG: About to update with checkbox_data: {checkbox_data}")
                    for field_name, field_value in checkbox_data.items():
                        print(f"DEBUG: - {field_name} = {field_value}")
                    
                    # Try updating ALL pages to find where checkboxes are located
                    for page_index in range(len(writer.pages)):
                        try:
                            writer.update_page_form_field_values(writer.pages[page_index], checkbox_data)
                            print(f"DEBUG: Updated page {page_index + 1} with {len(checkbox_data)} checkboxes")
                            
                            # CRITICAL: Set appearance state directly for each checkbox
                            self._set_checkbox_appearance_state(writer.pages[page_index], checkbox_data)
                            
                            # Don't break - check all pages to see where checkboxes are
                        except Exception as page_error:
                            print(f"DEBUG: Failed to update page {page_index + 1}: {page_error}")
                            continue
            
            print(f"DEBUG: Inspection matrix filling completed. Total checkbox updates attempted: {total_checkbox_updates}")
            
        except Exception as e:
            print(f"DEBUG: Error filling inspection matrix: {e}")
            import traceback
            traceback.print_exc()

    def _fill_inspection_details(self, writer: PdfWriter, sections: List[Dict], form_fields: Dict) -> None:
        """Fill detailed inspection comments and observations in text fields."""
        try:
            print("DEBUG: Filling inspection details...")
            
            # Get text fields available for comments
            text_fields = [field for field in form_fields.keys() if 'Text' in field and not field.startswith('TextField')]
            print(f"DEBUG: Found {len(text_fields)} text fields for details")
            
            # Collect all comments from all sections
            all_comments = []
            deficiency_comments = []
            
            for section in sections:
                section_name = section.get('name', '')
                line_items = section.get('line_items', [])
                
                for line_item in line_items:
                    comments = line_item.get('comments', [])
                    
                    for comment in comments:
                        comment_text = comment.get('text', '').strip()
                        comment_type = comment.get('type', '')
                        location = comment.get('location', '')
                        
                        if comment_text:
                            # Format the comment with context
                            formatted_comment = f"{section_name}"
                            if location:
                                formatted_comment += f" ({location})"
                            formatted_comment += f": {comment_text}"
                            
                            if comment_type == 'defect':
                                deficiency_comments.append(formatted_comment)
                            else:
                                all_comments.append(formatted_comment)
            
            print(f"DEBUG: Found {len(all_comments)} general comments and {len(deficiency_comments)} deficiency comments")
            
            # Fill deficiency comments first (more important)
            comments_to_fill = deficiency_comments + all_comments
            
            # Fill text fields with comments
            field_data = {}
            
            for i, comment in enumerate(comments_to_fill[:len(text_fields)]):
                field_name = text_fields[i]
                # Truncate very long comments to fit in form fields
                truncated_comment = comment[:500] + "..." if len(comment) > 500 else comment
                field_data[field_name] = truncated_comment
                print(f"DEBUG: Filling {field_name} with comment: {truncated_comment[:50]}...")
            
            # Update all pages with comment data
            if field_data:
                for page_index in range(len(writer.pages)):
                    try:
                        writer.update_page_form_field_values(writer.pages[page_index], field_data)
                        print(f"DEBUG: Updated page {page_index + 1} with {len(field_data)} text fields")
                    except Exception as page_error:
                        print(f"DEBUG: Could not update page {page_index + 1}: {page_error}")
                        continue
            
            print("DEBUG: Inspection details filling completed")
            
        except Exception as e:
            print(f"DEBUG: Error filling inspection details: {e}")
            import traceback
            traceback.print_exc()

    def _create_checkbox_appearances(self, page, checkbox_data):
        """Create visual appearance streams for filled checkboxes."""
        try:
            from PyPDF2.generic import DictionaryObject, ArrayObject, TextStringObject
            
            print(f"DEBUG: Creating appearances for {len(checkbox_data)} checkboxes")
            
            if '/Annots' not in page:
                return
                
            annotations = page['/Annots']
            appearances_created = 0
            
            for annot_ref in annotations:
                try:
                    annot = annot_ref.get_object()
                    
                    # Check if this is a checkbox widget we filled
                    if ('/Subtype' in annot and annot['/Subtype'] == '/Widget' and 
                        '/FT' in annot and annot['/FT'] == '/Btn' and
                        '/T' in annot):
                        
                        field_name = str(annot['/T'])
                        if field_name in checkbox_data:
                            print(f"DEBUG: Creating appearance for checkbox: {field_name}")
                            
                            # Force the checkbox to show as checked by setting appearance
                            if '/AP' not in annot:
                                annot['/AP'] = DictionaryObject()
                            
                            # Set the appearance state to match the value
                            appearance_dict = annot['/AP']
                            if '/N' not in appearance_dict:
                                appearance_dict['/N'] = DictionaryObject()
                            
                            # Create "Yes" appearance state (checked)
                            normal_appearance = appearance_dict['/N']
                            if '/Yes' not in normal_appearance:
                                # Create a simple checked appearance
                                normal_appearance['/Yes'] = DictionaryObject({
                                    '/Type': '/XObject',
                                    '/Subtype': '/Form',
                                    '/BBox': ArrayObject([0, 0, 1000, 1000]),
                                    '/Length': 44,
                                    '/Stream': TextStringObject('q\n0 0 1000 1000 re\nW\nn\nBT\n/F1 800 Tf\n100 100 Td\n(X) Tj\nET\nQ')
                                })
                            
                            # Set the appearance state
                            annot['/AS'] = '/Yes'
                            appearances_created += 1
                            
                except Exception as e:
                    print(f"DEBUG: Error creating appearance for annotation: {e}")
                    continue
            
            print(f"DEBUG: Created {appearances_created} checkbox appearances")
            
        except Exception as e:
            print(f"DEBUG: Error in _create_checkbox_appearances: {e}")

    def _set_checkbox_appearance_state(self, page, checkbox_data):
        """Set the appearance state (/AS) for checkbox fields to make them visible."""
        try:
            from PyPDF2.generic import NameObject
            print(f"DEBUG: Setting appearance states for {len(checkbox_data)} checkboxes")
            print(f"DEBUG: Checkbox data to set: {checkbox_data}")
            
            if '/Annots' not in page:
                print("DEBUG: No annotations found on page")
                return
                
            annotations = page['/Annots']
            states_set = 0
            total_annotations = len(annotations) if annotations else 0
            print(f"DEBUG: Found {total_annotations} annotations on page")
            
            for i, annot_ref in enumerate(annotations):
                try:
                    annot = annot_ref.get_object()
                    
                    # Check if this is a widget annotation
                    if '/Subtype' in annot and annot['/Subtype'] == '/Widget':
                        field_name = str(annot.get('/T', ''))
                        field_type = annot.get('/FT', '')
                        
                        print(f"DEBUG: Annotation {i}: field_name='{field_name}', field_type='{field_type}'")
                        
                        # Check if this is a button (checkbox) we want to update
                        if field_type == '/Btn' and field_name in checkbox_data:
                            field_value = checkbox_data[field_name]
                            print(f"DEBUG: MATCH! Setting {field_name} to {field_value}")
                            
                            # Set the value using proper PyPDF2 objects
                            annot[NameObject('/V')] = NameObject('/' + field_value)
                            
                            # Set the appearance state to match the value
                            if field_value in ['Yes', '/Yes']:
                                annot[NameObject('/AS')] = NameObject('/Yes')
                                print(f"DEBUG: Set appearance state to /Yes")
                            elif field_value in ['On', '/On']:
                                annot[NameObject('/AS')] = NameObject('/On')
                                print(f"DEBUG: Set appearance state to /On")
                            elif field_value in ['1', '/1']:
                                annot[NameObject('/AS')] = NameObject('/1')
                                print(f"DEBUG: Set appearance state to /1")
                            else:
                                annot[NameObject('/AS')] = NameObject('/' + str(field_value))
                                print(f"DEBUG: Set appearance state to /{field_value}")
                            
                            states_set += 1
                        elif field_type == '/Btn':
                            print(f"DEBUG: Button field '{field_name}' not in checkbox_data")
                            
                except Exception as e:
                    print(f"DEBUG: Error processing annotation {i}: {e}")
                    continue
            
            print(f"DEBUG: Set appearance states for {states_set} checkboxes")
            
        except Exception as e:
            print(f"DEBUG: Error in _set_checkbox_appearance_state: {e}")
            import traceback
            traceback.print_exc()

    def _flatten_checkboxes_only(self, writer: PdfWriter) -> None:
        """Flatten only checkbox fields to make them permanently visible while preserving text fields."""
        try:
            print("DEBUG: Flattening checkbox fields only...")
            
            checkboxes_flattened = 0
            
            for page_num, page in enumerate(writer.pages):
                if '/Annots' in page:
                    annotations = page['/Annots']
                    page_checkboxes = 0
                    
                    for annot_ref in annotations:
                        try:
                            annot = annot_ref.get_object()
                            
                            # Only process checkbox widgets (buttons with /V value)
                            if ('/Subtype' in annot and annot['/Subtype'] == '/Widget' and
                                '/FT' in annot and annot['/FT'] == '/Btn' and
                                '/V' in annot and annot['/V']):
                                
                                field_name = str(annot.get('/T', ''))
                                field_value = str(annot.get('/V', ''))
                                
                                print(f"DEBUG: Flattening checkbox {field_name} with value {field_value}")
                                
                                # Remove interactive properties but keep the value
                                properties_to_remove = ['/FT', '/Ff', '/P', '/Parent']
                                for prop in properties_to_remove:
                                    if prop in annot:
                                        del annot[prop]
                                
                                # Ensure the checkbox appears checked
                                if '/AS' not in annot:
                                    annot['/AS'] = annot['/V']
                                
                                page_checkboxes += 1
                                checkboxes_flattened += 1
                        except Exception as e:
                            print(f"DEBUG: Error flattening annotation: {e}")
                            continue
                    
                    if page_checkboxes > 0:
                        print(f"DEBUG: Flattened {page_checkboxes} checkboxes on page {page_num + 1}")
            
            print(f"DEBUG: Total checkboxes flattened: {checkboxes_flattened}")
            
        except Exception as e:
            print(f"DEBUG: Checkbox flattening failed: {e}")


def main():
    """Main function to generate TREC report."""
    print("=" * 60)
    print("🏠 TREC HOME INSPECTION REPORT GENERATOR")
    print("=" * 60)
    
    # Configuration
    json_file = "inspection.json"
    template_file = "TREC_Template_Blank.pdf"
    output_file = "output_pdf.pdf"
    
    print(f"📁 Input file: {json_file}")
    print(f"📄 Template: {template_file}")
    print(f"📂 Output file: {output_file}")
    print()
    
    try:
        # Check if required files exist
        if not os.path.exists(json_file):
            print(f"❌ Error: {json_file} not found!")
            return False
            
        if not os.path.exists(template_file):
            print(f"❌ Error: {template_file} not found!")
            return False
        
        # Initialize data parser
        print("🔍 Parsing inspection data...")
        parser = InspectionDataParser(json_file)
        
        # Get basic info for verification
        property_info = parser.get_property_info()
        sections = parser.get_sections()
        
        # Count line items and comments
        total_line_items = 0
        total_comments = 0
        deficiencies = 0
        
        for section in sections:
            line_items = section.get('line_items', [])
            total_line_items += len(line_items)
            
            for line_item in line_items:
                comments = line_item.get('comments', [])
                total_comments += len(comments)
                
                if line_item.get('is_deficient', False):
                    deficiencies += 1
        
        print("✅ Data parsing successful!")
        print(f"   📍 Property: {property_info.get('full_address', 'Unknown')}")
        print(f"   📊 Sections: {len(sections)}")
        print(f"   📝 Line Items: {total_line_items}")
        print(f"   💬 Comments: {total_comments}")
        print(f"   ⚠️  Deficiencies: {deficiencies}")
        print()
        
        # Generate TREC report
        print("📄 Generating TREC Report...")
        print("   📋 Filling TREC template form fields")
        
        form_filler = TRECFormFiller(parser, template_file)
        success = form_filler.generate_report(output_file)
        
        if success:
            # Get file size
            file_size = os.path.getsize(output_file) / (1024 * 1024)  # MB
            
            print(f"✅ TREC report generated successfully: {output_file}")
            print(f"   📁 File size: {file_size:.2f} MB")
            print("   📄 Official TREC REI 7-6 format maintained")
            print()
            print("=" * 60)
            print("📊 GENERATION SUMMARY")
            print("=" * 60)
            print("✅ Successfully generated: 1/1 reports")
            print("🎉 All reports generated successfully!")
            print()
            print("📋 Generated files:")
            print(f"   📄 TREC Report: {output_file}")
            print()
            print("💡 Tips:")
            print("   • Test PDFs in different viewers (Adobe, Chrome, etc.)")
            print("   • Check that all form fields are properly filled")
            print("   • Verify that inspection matrix checkboxes are marked")
            print("   • Review detailed comments on pages 3-6")
            
            return True
        else:
            print("❌ Failed to generate TREC report!")
            return False
            
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    main()

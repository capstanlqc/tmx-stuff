#!/usr/bin/env python3
import sys
import os
import xml.etree.ElementTree as ET
import re

def get_cleaned_output_path(input_path):
    """Set output path to 'cleaned/<input_filename>'"""
    dir_name = os.path.dirname(input_path)
    base_name = os.path.basename(input_path)
    cleaned_dir = os.path.join(dir_name, 'cleaned')
    if not os.path.exists(cleaned_dir):
        os.makedirs(cleaned_dir)
    return os.path.join(cleaned_dir, base_name)

def is_number(s):
    """Check if string can be converted to a number"""
    try:
        float(s)
        return True
    except (ValueError, TypeError):
        return False

def is_only_tags(s):
    """Check if string contains only XML tags"""
    if not s:
        return False
    return re.sub(r'<[^>]+>', '', s).strip() == ''

def is_number_range_or_group(s):
    """Check if string is a number range (e.g., '1-5', '3 ~ 7') or group of numbers (e.g., '1 2 3', '4,5,6')"""
    if not s:
        return False
    s = s.strip()
    # Remove tags before checking
    s = re.sub(r'<[^>]+>', '', s)
    # Number range: 1-5, 1~5, 1 – 5, etc.
    if re.match(r'^\s*\d+\s*[-–~]\s*\d+\s*$', s):
        return True
    # Group of numbers: 1 2 3, 4,5,6, etc.
    if re.match(r'^\s*(\d+[\s,;])+(\d+)\s*$', s):
        return True
    return False

def contains_letter(s):
    """Check if string contains at least one letter (any language)"""
    if not s:
        return False
    # Remove tags before checking
    s = re.sub(r'<[^>]+>', '', s)
    return bool(re.search(r'\p{L}', s, re.UNICODE)) or bool(re.search(r'[A-Za-z]', s))

def process_tmx(input_path):
    """Process TMX file and remove invalid translation units"""
    tree = ET.parse(input_path)
    root = tree.getroot()
    
    # TMX namespace handling
    ns = {'tmx': 'http://www.lisa.org/tmx14'}
    
    # Find all translation units
    tus = root.findall('.//tmx:tu', ns)
    valid_tus = []

    for tu in tus:
        tuvs = tu.findall('tmx:tuv', ns)
        num_tuvs = len(tuvs)
        skip_tu = False

        seg_texts = []
        for tuv in tuvs:
            seg = tuv.find('tmx:seg', ns)
            seg_text = seg.text if seg is not None else None
            seg_texts.append(seg_text)

            # Remove if segment is missing or empty
            if not seg_text or not seg_text.strip():
                skip_tu = True
                break

            # Remove if segment is only tags, only numbers, number range/group, or lacks any letters
            if (is_only_tags(seg_text) or
                is_number(seg_text) or
                is_number_range_or_group(seg_text) or
                not contains_letter(seg_text)):
                skip_tu = True
                break

        if skip_tu:
            continue

        # Remove if both TUVs are present and identical
        if num_tuvs == 2 and seg_texts[0] == seg_texts[1]:
            continue

        valid_tus.append(tu)

    # Create new TMX structure
    new_root = ET.Element(root.tag, root.attrib)
    for child in root:
        if child.tag != f'{{{ns["tmx"]}}}body':
            new_root.append(child)
    
    new_body = ET.SubElement(new_root, f'{{{ns["tmx"]}}}body')
    for tu in valid_tus:
        new_body.append(tu)
    
    # Write output file
    base_name, ext = os.path.splitext(input_path)
    output_path = get_cleaned_output_path(input_path)
    ET.ElementTree(new_root).write(
        output_path, 
        encoding='utf-8', 
        xml_declaration=True
    )
    return output_path

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Usage: tmx_cleaner.py <tmx_file>")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = process_tmx(input_file)
    print(f"Cleaned TMX saved to: {output_file}")

import os
import re
import argparse
import pandas as pd
from collections import defaultdict, Counter
from xml.etree.ElementTree import Element, SubElement, tostring, Comment
from pathlib import Path
import xml.dom.minidom

def extract_data(file_path, source_col, target_col, sheet_pattern, alttype):
    data = []
    xls = pd.ExcelFile(file_path)
    for sheet_name in xls.sheet_names:
        if not re.match(sheet_pattern, sheet_name):
            continue
        df = xls.parse(sheet_name, header=1, keep_default_na=False, na_values=[])
        required_cols = ["Segment ID", source_col, target_col]
        if not all(col in df.columns for col in required_cols):
            continue
        has_altuniq = 'Alt/Uniq' in df.columns
        nrows = len(df)
        for idx, row in df.iterrows():
            target_text = row[target_col]
            alt_uniq = row['Alt/Uniq'] if has_altuniq else None
            is_forced_alt = alt_uniq and 'a' in str(alt_uniq).lower()
            # Only collect if:
            # - forced alternative (even if target_text is None/empty)
            # - OR target_text is not null/empty
            if not is_forced_alt:
                # skip if target_text is None, empty, or only whitespace
                if pd.isnull(target_text) or str(target_text).strip() == '':
                    continue
            source_text = row[source_col]
            segment_id = row["Segment ID"] if str(row["Segment ID"]).strip() else None
            prev_source = df.iloc[idx-1][source_col] if idx > 0 else ''
            next_source = df.iloc[idx+1][source_col] if idx < nrows-1 else ''
            if alttype == 'context' or not segment_id:
                data.append({
                    "source_text": source_text,
                    "target_text": target_text,
                    "prev_source": prev_source,
                    "next_source": next_source,
                    "segment_id": None,
                    "alt_uniq": alt_uniq
                })
            else:
                data.append({
                    "source_text": source_text,
                    "target_text": target_text,
                    "segment_id": str(segment_id),
                    "prev_source": None,
                    "next_source": None,
                    "alt_uniq": alt_uniq
                })
    return data

def categorize_translations(data, alttype):
    from collections import defaultdict, Counter
    grouped = defaultdict(list)
    for item in data:
        grouped[item["source_text"]].append(item)

    default_translations = []
    alternative_translations = []

    for source, items in grouped.items():
        # Separate forced and non-forced items
        forced_items = []
        nonforced_items = []
        for item in items:
            is_forced = item.get('alt_uniq') and 'a' in str(item['alt_uniq']).lower()
            if is_forced:
                # Forced alternative: add only to alternative_translations
                alternative_translations.append(item)
            else:
                nonforced_items.append(item)

        # Now process non-forced items for default/alternative logic
        if nonforced_items:
            target_counter = Counter(item["target_text"] for item in nonforced_items)
            unique_targets = list(target_counter.keys())

            if len(unique_targets) == 1:
                default_translations.append({
                    "source_text": source,
                    "target_text": unique_targets[0]
                })
            else:
                max_count = max(target_counter.values())
                max_targets = [t for t, count in target_counter.items() if count == max_count]
                # Pick the first encountered as default
                default_target = None
                for item in nonforced_items:
                    if item["target_text"] in max_targets:
                        default_target = item["target_text"]
                        break
                if default_target is not None:
                    default_translations.append({
                        "source_text": source,
                        "target_text": default_target
                    })
                # Add all non-forced items for other targets to alternatives
                for variant_target in unique_targets:
                    if variant_target == default_target:
                        continue
                    for item in nonforced_items:
                        if item["target_text"] == variant_target:
                            alternative_translations.append(item)
        # If there are only forced items (no non-forced), do not add to default

    # Deduplicate alternative translations
    seen = set()
    filtered_alternatives = []
    for item in alternative_translations:
        if alttype == 'context' or (not item.get('segment_id')):
            key = (
                item["source_text"], item["target_text"],
                item.get("prev_source"), item.get("next_source")
            )
        else:
            key = (
                item["source_text"], item["target_text"], item.get("segment_id")
            )
        if key not in seen:
            seen.add(key)
            filtered_alternatives.append(item)
    return {
        "default_translations": default_translations,
        "alternative_translations": filtered_alternatives
    }

def ensure_output_dir_exists(output_path):
    output_dir = os.path.dirname(output_path)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

def create_tmx(data_dict, output_path, source_lang, target_lang, input_file_path, alttype):
    tmx = Element('tmx', version='1.4')
    header = SubElement(tmx, 'header',
                        creationtool='Excel2TMX.py',
                        creationtoolversion='1.0',
                        segtype='sentence',
                        adminlang='en-us',
                        srclang=source_lang,
                        datatype='PlainText')

    body = SubElement(tmx, 'body')

    filename = os.path.basename(input_file_path)

    # Default translations first
    for item in data_dict['default_translations']:
        tu = SubElement(body, 'tu')
        tuv_source = SubElement(tu, 'tuv', {'xml:lang': source_lang})
        seg_source = SubElement(tuv_source, 'seg')
        seg_source.text = str(item['source_text']) if item['source_text'] is not None else ''

        tuv_target = SubElement(tu, 'tuv', {'xml:lang': target_lang})
        seg_target = SubElement(tuv_target, 'seg')
        seg_target.text = str(item['target_text']) if item['target_text'] is not None else ''

    # Insert XML comment to separate groups
    body.append(Comment('Alternative translations'))

    # Alternative translations after the comment
    for item in data_dict['alternative_translations']:
        tu = SubElement(body, 'tu')
        SubElement(tu, 'prop', {'type': 'file'}).text = filename
        if alttype == 'id' and item.get('segment_id'):
            SubElement(tu, 'prop', {'type': 'id'}).text = str(item['segment_id'])
        else:
            SubElement(tu, 'prop', {'type': 'prev'}).text = str(item.get('prev_source') or '')
            SubElement(tu, 'prop', {'type': 'next'}).text = str(item.get('next_source') or '')

        tuv_source = SubElement(tu, 'tuv', {'xml:lang': source_lang})
        seg_source = SubElement(tuv_source, 'seg')
        seg_source.text = str(item['source_text']) if item['source_text'] is not None else ''

        tuv_target = SubElement(tu, 'tuv', {'xml:lang': target_lang})
        seg_target = SubElement(tuv_target, 'seg')
        # For forced alternatives, allow empty/null target
        if item.get('alt_uniq') and 'a' in str(item.get('alt_uniq')).lower():
            seg_target.text = '' if not item['target_text'] else str(item['target_text'])
        else:
            seg_target.text = str(item['target_text']) if item['target_text'] else ''

    rough_string = tostring(tmx, 'utf-8')
    reparsed = xml.dom.minidom.parseString(rough_string)
    pretty_xml = reparsed.toprettyxml(indent="  ")

    # Remove lines containing only whitespace
    cleaned_xml = "\n".join([line for line in pretty_xml.splitlines() if line.strip()])

    ensure_output_dir_exists(output_path)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(cleaned_xml)

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('file_path', help='Path to XLSX file')
    parser.add_argument('--sl', required=True, help='Source language column name (BCP47)')
    parser.add_argument('--tl', required=True, help='Target language column name (BCP47)')
    parser.add_argument('--sheet-pattern', default='.*', help='Regex pattern to match sheet names')
    parser.add_argument('--alttype', default='id', choices=['id', 'context'], help='Alternative translation grouping: "id" or "context"')
    parser.add_argument('--omt', action='store_true', help='If set, output TMX to ../tm/excel2tmx/ relative to input file; otherwise to ../excel2tmx_output/')
    args = parser.parse_args()

    source_lang = args.sl
    target_lang = args.tl
    alttype = args.alttype

    data = extract_data(args.file_path, source_lang, target_lang, args.sheet_pattern, alttype)
    categorized = categorize_translations(data, alttype)
    print(categorized)

    input_path = Path(args.file_path).resolve()
    input_dir = input_path.parent
    if args.omt:
        output_dir = (input_dir / '..' / 'tm' / 'excel2tmx').resolve()
    else:
        output_dir = (input_dir / '..' / 'excel2tmx_output').resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    base_name = input_path.stem
    output_path = output_dir / f"{base_name}.tmx"

    if categorized['default_translations'] or categorized['alternative_translations']:
        create_tmx(categorized, output_path, source_lang, target_lang, args.file_path, alttype)
        print(f"TMX file created at: {output_path}")
        print(f"Default translations: {len(categorized['default_translations'])}")
        print(f"Alternative translations: {len(categorized['alternative_translations'])}")
    else:
        print(f"No segments collected from {args.file_path}; TMX file not created.")

if __name__ == '__main__':
    main()

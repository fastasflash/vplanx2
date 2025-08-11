#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
qvm2vplanx_full.py

Conversion script for QVM Excel plans to Cadence vPlanx XML.
Parses Title, Description, Link, and Type from XLSX using only Python stdlib,
and generates a valid .vplanx without external dependencies.
"""
import sys
import os
import re
import uuid
import time
import logging
import argparse
import xml.etree.ElementTree as ET
import gzip
import zipfile
import xml.etree.ElementTree as xmlET

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger('qvm2vplanx')

version = '1.0.4'

def parse_excel(xlsx_path):
    if not zipfile.is_zipfile(xlsx_path):
        raise Exception(f"File is not a valid XLSX archive: {xlsx_path}")

    with zipfile.ZipFile(xlsx_path, 'r') as z:
        try:
            data = z.read('xl/sharedStrings.xml')
            ss_root = xmlET.fromstring(data)
            shared_strings = [si.find('.//{*}t').text.strip() if si.find('.//{*}t') is not None else '' for si in ss_root.findall('{*}si')]
        except KeyError:
            shared_strings = []

        sheet_files = [n for n in z.namelist() if re.match(r'xl/worksheets/sheet\d+\.xml', n)]
        if not sheet_files:
            raise Exception('No worksheet found in XLSX archive.')

        data = z.read(sheet_files[0])

    root = xmlET.fromstring(data)
    rows = []
    for row in root.findall('.//{*}row'):
        values = []
        for c in row.findall('{*}c'):
            v = c.find('{*}v')
            if v is not None:
                idx = int(v.text) if c.get('t') == 's' else None
                values.append(shared_strings[idx] if idx is not None and idx < len(shared_strings) else v.text)
            else:
                values.append('')
        rows.append(values)

    header_idx, headers = None, {}
    for i, row in enumerate(rows[:20]):
        low_row = [str(cell).strip().lower() for cell in row]
        if 'title' in low_row and 'link' in low_row:
            header_idx = i
            headers = {key: idx for idx, key in enumerate(low_row)}
            break

    if header_idx is None:
        logger.error("Could not find 'Title' and 'Link' headers in the first 20 rows.")
        logger.debug(f"Rows scanned: {rows[:20]}")
        raise Exception('Header row not found containing Title and Link columns.')

    entries = []
    for row in rows[header_idx+1:]:
        if not any(row):
            continue
        title = row[headers['title']] if headers.get('title') is not None and headers['title'] < len(row) else ''
        if not title:
            continue
        desc = row[headers['description']] if 'description' in headers and headers['description'] < len(row) else ''
        link = row[headers['link']] if 'link' in headers and headers['link'] < len(row) else ''
        type_ = row[headers['type']] if 'type' in headers and headers['type'] < len(row) else ''

        entries.append({'title': title.strip(), 'desc': desc.strip(), 'link': link.strip(), 'type': type_.strip()})

    return entries

def build_vplanx(entries, plan_name):
    ET.register_namespace('vplanx', 'http://www.cadence.com/vplanx')
    root = ET.Element('vplanx:plan', {'xmlns:vplanx': 'http://www.cadence.com/vplanx'})

    md = ET.SubElement(root, 'metaData', {'id': str(uuid.uuid1())})
    ET.SubElement(md, 'name').text = plan_name
    ET.SubElement(md, 'planId').text = md.get('id')
    ET.SubElement(md, 'sourceTool').text = 'qvm2vplanx'
    ET.SubElement(md, 'toolVersion').text = version
    ET.SubElement(md, 'schemaVersion').text = '1.0'
    ET.SubElement(md, 'buildTime').text = time.strftime('%Y-%m-%d %H:%M:%S')

    root_elements = ET.SubElement(root, 'rootElements')

    for entry in entries:
        sect = ET.SubElement(root_elements, 'section', {'id': str(uuid.uuid1())})
        ET.SubElement(sect, 'name').text = entry['title']

        attrs = ET.SubElement(sect, 'attributes')
        for attr_name, attr_value in [('details', entry['desc']), ('type', entry['type']), ('planned_elements', '1')]:
            attr = ET.SubElement(attrs, 'attribute')
            ET.SubElement(attr, 'name').text = attr_name
            ET.SubElement(attr, 'value').text = attr_value

        mp = ET.SubElement(sect, 'metricsPort', {'id': str(uuid.uuid1())})
        ET.SubElement(mp, 'name').text = entry['title']

        mapping_patterns = ET.SubElement(mp, 'mappingPatterns')
        mapping_pattern = ET.SubElement(mapping_patterns, 'mappingPattern', {'id': str(uuid.uuid1())})
        domain_el = ET.SubElement(mapping_pattern, 'domains')
        ET.SubElement(domain_el, 'domain').text = 'HDL'
        entity_el = ET.SubElement(mapping_pattern, 'entityKinds')
        ET.SubElement(entity_el, 'entityKind').text = 'INSTANCE'
        ET.SubElement(mapping_pattern, 'pattern').text = entry['link']

    return ET.ElementTree(root)

def save_tree(tree, filename, gzip_out=True):
    with gzip.open(filename, 'wb') if gzip_out else open(filename, 'wb') as f:
        tree.write(f, encoding='utf-8', xml_declaration=True)

def main(args):
    parser = argparse.ArgumentParser(description='Convert XLSX to vplanx')
    parser.add_argument('xlsx', help='Input XLSX file')
    parser.add_argument('--out', help='Output .vplanx file name')
    parser.add_argument('--no-gzip', action='store_true', help='Do not gzip output')

    parsed = parser.parse_args(args)
    plan_name = os.path.splitext(os.path.basename(parsed.xlsx))[0]
    entries = parse_excel(parsed.xlsx)
    tree = build_vplanx(entries, plan_name)

    outname = parsed.out or (plan_name + '.vplanx')
    save_tree(tree, outname, not parsed.no_gzip)
    print('Created', outname)

if __name__ == '__main__':
    main(sys.argv[1:])

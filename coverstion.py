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

# Logger
logger = logging.getLogger('qvm2vplanx')
logging.basicConfig(level=logging.INFO)

# Version
version = '1.0.2'

# ----------------------------------------------------------------------------
# Core conversion logic
# ----------------------------------------------------------------------------

def parse_excel(xlsx_path):
    """
    Parse the .xlsx file for entries using Python stdlib (zipfile + XML).
    Extracts rows with columns Title, Description, Link, and Type.
    """
    if not zipfile.is_zipfile(xlsx_path):
        raise Exception(f"File is not a valid XLSX archive: {xlsx_path}")
    with zipfile.ZipFile(xlsx_path, 'r') as z:
        # Read sharedStrings
        ss = []
        try:
            data = z.read('xl/sharedStrings.xml')
            root = xmlET.fromstring(data)
            for si in root.findall('{*}si'):
                t = si.find('.//{*}t')
                ss.append(t.text or '')
        except KeyError:
            ss = []
        # Read first worksheet (dynamically pick sheetN.xml)
        sheet_files = [n for n in z.namelist() if n.startswith('xl/worksheets/sheet') and n.endswith('.xml')]
        if not sheet_files:
            raise Exception('No sheet*.xml found in XLSX archive')
        data = z.read(sheet_files[0])
    root = xmlET.fromstring(data)
    rows = []
    for row in root.findall('.//{*}row'):
        values = []
        for c in row.findall('{*}c'):
            v = c.find('{*}v')
            if v is None:
                values.append('')
            else:
                if c.get('t') == 's':
                    idx = int(v.text)
                    values.append(ss[idx] if idx < len(ss) else '')
                else:
                    values.append(v.text or '')
        rows.append(values)
    # locate header row (first 10 rows)
    header_idx = None
    headers = {}
    for i, row in enumerate(rows[:10]):
        low = [str(v).strip().lower() for v in row]
        if 'title' in low and 'link' in low:
            header_idx = i
            for j, name in enumerate(low):
                if name:
                    headers[name] = j
            break
    if header_idx is None:
        raise Exception('Header row not found containing Title and Link')
    # parse entries
    entries = []
    for row in rows[header_idx+1:]:
        if len(row) <= headers.get('title', -1):
            continue
        title = row[headers['title']]
        if not title:
            continue
        desc = row[headers.get('description', '')] if 'description' in headers else ''
        link = row[headers['link']]
        type_ = row[headers.get('type', '')] if 'type' in headers else ''
        entries.append({'title': title, 'desc': desc or '', 'link': link or '', 'type': type_ or ''})
    return entries


def build_vplanx(entries, plan_name):
    ET.register_namespace('vplanx', 'http://www.cadence.com/vplanx')
    root = ET.Element('vplanx:plan', {'xmlns:vplanx': 'http://www.cadence.com/vplanx'})
    # metadata
    md = ET.SubElement(root, 'metaData', {'id': str(uuid.uuid1())})
    ET.SubElement(md, 'name').text = plan_name
    ET.SubElement(md, 'planId').text = md.get('id')
    ET.SubElement(md, 'sourceTool').text = 'qvm2vplanx'
    ET.SubElement(md, 'toolVersion').text = version
    ET.SubElement(md, 'schemaVersion').text = '1.0'
    ET.SubElement(md, 'buildTime').text = time.strftime('%Y-%m-%d %H:%M:%S')
    # rootElements
    root_elements = ET.SubElement(root, 'rootElements')
    # sections
    for e in entries:
        sect = ET.SubElement(root_elements, 'section', {'id': str(uuid.uuid1())})
        ET.SubElement(sect, 'name').text = str(e['title'])
        # attributes
        attrs = ET.SubElement(sect, 'attributes')
        for key, val in [('details', e['desc']), ('type', e['type']), ('planned_elements', '1')]:
            a = ET.SubElement(attrs, 'attribute')
            ET.SubElement(a, 'name').text = key
            ET.SubElement(a, 'value').text = str(val)
        # metricsPort
        mp = ET.SubElement(sect, 'metricsPort', {'id': str(uuid.uuid1())})
        ET.SubElement(mp, 'name').text = str(e['title'])
        mps = ET.SubElement(mp, 'mappingPatterns')
        mp_node = ET.SubElement(mps, 'mappingPattern', {'id': str(uuid.uuid1())})
        # domains
        dms = ET.SubElement(mp_node, 'domains')
        ET.SubElement(dms, 'domain').text = 'HDL'
        # entityKinds
        ems = ET.SubElement(mp_node, 'entityKinds')
        ET.SubElement(ems, 'entityKind').text = 'INSTANCE'
        # pattern
        pat = ET.SubElement(mp_node, 'pattern')
        pat.text = str(e['link'])
    return ET.ElementTree(root)


def save_tree(tree, filename, gzip_out=True):
    if gzip_out:
        with gzip.open(filename, 'wb') as f:
            tree.write(f, encoding='utf-8', xml_declaration=True)
    else:
        tree.write(filename, encoding='utf-8', xml_declaration=True)

# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------

def main(args):
    parser = argparse.ArgumentParser(description='Convert XLSX to vplanx')
    parser.add_argument('xlsx', help='Input XLSX file')
    parser.add_argument('--out', help='Output .vplanx file name')
    parser.add_argument('--no-gzip', action='store_true', help='Do not gzip output')
    parsed = parser.parse_args(args)
    xlsx = parsed.xlsx
    if not os.path.exists(xlsx):
        print('Error: file not found', xlsx)
        sys.exit(1)
    plan_name = os.path.splitext(os.path.basename(xlsx))[0]
    entries = parse_excel(xlsx)
    tree = build_vplanx(entries, plan_name)
    outname = parsed.out or (plan_name + '.vplanx')
    save_tree(tree, outname, not parsed.no_gzip)
    print('Created', outname)

if __name__ == '__main__':
    main(sys.argv[1:])

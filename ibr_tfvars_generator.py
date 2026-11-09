#!/usr/bin/env python3
"""Terraform Script Generator — dynamic CIQ converter with verified horizontal and vertical target-block extraction. Python 3.10+, stdlib only.
Output is generic .tfvars HCL maps/lists. Adapt root variable names to your Terraform module.
"""
from __future__ import annotations
import argparse, csv, datetime as dt, fnmatch, json, re, sys, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from collections import Counter

NS={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships','p':'http://schemas.openxmlformats.org/package/2006/relationships'}
NAME=re.compile(r'[^a-z0-9_]+')
SKIP={'','see value below.','see value below','[enc]','[tbd]','[telus-input]','#ref!','#n/a','n/a','skip','✓','◌'}
DEFAULT=re.compile(r'^\[\s*default\s*=\s*(.*?)\s*\]$',re.I)
CELL=re.compile(r'([A-Z]+)(\d+)')

def colnum(label):
 n=0
 for c in label:n=n*26+ord(c)-64
 return n

def clean_name(v):
 """Normalize CIQ/YANG leaf paths consistently, including wrapped Excel cells.

 A displayed child path may contain a newline or spaces surrounding its slash,
 e.g. `tal-info-ta-entry/\n tal-info-mcc`. Both metadata declarations and
 the target-table headings must resolve to the same canonical key.
 """
 v=re.sub(r'\s*/\s*', '/', str(v).strip())
 v=v.replace('/','_').replace('-','_').lower()
 v=NAME.sub('_',v).strip('_')
 return ('v_'+v) if v and v[0].isdigit() else v

class DesignDefault(str):
 """A target cell explicitly marked [default = ...], distinct from an assigned value."""

def normalize(v,omit_markers=True):
 if v is None:return None
 v=str(v).replace('\xa0',' ').strip()
 if v.lower() in SKIP:return None
 m=DEFAULT.match(v)
 if m:
  value=m.group(1).strip()
  return DesignDefault(value) if value else None
 if not v:return None
 if v.startswith('{') and v.endswith('}'):
  parts=[p.strip().strip('"\'') for p in v[1:-1].split(',') if p.strip()]
  return [normalize(p,omit_markers) for p in parts if normalize(p,omit_markers) is not None]
 return v

INTEGER = re.compile(r'^[+-]?(?:0|[1-9][0-9]*)$')
DECIMAL = re.compile(r'^[+-]?[0-9]+\.[0-9]+$')
LEADING_ZERO = re.compile(r'^[+-]?0[0-9]+$')

def hcl(v):
 """Emit whole numbers as HCL numbers, but quote decimals and leading-zero identifiers.

 Examples: 95 -> 95, 95.0 -> "95.0", 0040 -> "0040",
 {0, 1, 2} (normalized to list) -> [0, 1, 2].
 """
 if isinstance(v,list):return '['+', '.join(hcl(x) for x in v)+']'
 if v is None:return 'null'
 raw=str(v)
 if LEADING_ZERO.fullmatch(raw):return json.dumps(raw,ensure_ascii=False)
 if INTEGER.fullmatch(raw):return raw
 return json.dumps(raw,ensure_ascii=False)

def read_workbook(path):
 with zipfile.ZipFile(path) as z:
  shared=[]
  if 'xl/sharedStrings.xml' in z.namelist():
   root=ET.fromstring(z.read('xl/sharedStrings.xml'))
   shared=[''.join(n.itertext()) for n in root.findall('m:si',NS)]
  book=ET.fromstring(z.read('xl/workbook.xml'))
  rels=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
  relmap={r.attrib['Id']:r.attrib['Target'].lstrip('/') for r in rels}
  for s in book.findall('m:sheets/m:sheet',NS):
   sheet_name=s.attrib['name']; target=relmap[s.attrib['{%s}id'%NS['r']]]
   filename=target if target.startswith('xl/') else 'xl/'+target
   if filename not in z.namelist():continue
   rows=[]
   with z.open(filename) as source:
    for _,el in ET.iterparse(source,events=('end',)):
     if el.tag!='{%s}row'%NS['m']:continue
     row={}
     for cell in el.findall('m:c',NS):
      ref=cell.attrib.get('r','')
      match=CELL.match(ref)
      if not match:continue
      col=colnum(match.group(1))
      typ=cell.attrib.get('t')
      if typ=='inlineStr': val=''.join(cell.find('m:is',NS).itertext()) if cell.find('m:is',NS) is not None else ''
      else:
       v=cell.find('m:v',NS)
       if v is None:continue
       val=v.text or ''
       if typ=='s':val=shared[int(val)]
       elif typ=='b':val='true' if val=='1' else 'false'
       elif typ not in ('str','e') and val:
        # Keep numeric cells using Excel formatted display when zero masks apply.
        pass
      if val is not None and str(val).strip():row[col]=str(val).strip()
     if row:rows.append((int(el.attrib.get('r','0')),row))
     el.clear()
   yield sheet_name,rows

def is_parameter_name(v):
 v=re.sub(r'\s*/\s*','/',str(v).strip())
 # Allow any new YANG-style identifier, not just the currently known parameter names.
 return bool(re.fullmatch(r'[A-Za-z][A-Za-z0-9_.:-]*(?:/[A-Za-z][A-Za-z0-9_.:-]*)*',v)) and (('-' in v) or ('_' in v) or ('/' in v)) and v.lower() not in ('target-values','yang-path')

def is_header(row):
 vals=[v for c,v in row.items() if 2<=c<=120]
 keys=[v for v in vals if is_parameter_name(v) and ('-' in v or '/' in v or v.count('_')>=2)]
 return len(keys)>=2 and len(keys)>=max(2,min(4,len(vals)//2))

def take_table(rows,start,end,header_row,head_columns):
 records=[]; warns=[]
 cols=[(c,clean_name(v)) for c,v in sorted(header_row.items()) if c in head_columns and clean_name(v)]
 seen=set(); cols=[(c,n) for c,n in cols if not (n in seen or seen.add(n))]
 for rn,r in rows:
  if rn<=start or rn>=end:continue
  if is_header(r) or is_post_target_note(r):break
  raw=[r.get(c) for c,_ in cols]
  if not any(normalize(x) is not None for x in raw):continue
  rec={name:normalize(r.get(c)) for c,name in cols if normalize(r.get(c)) is not None}
  if not rec:continue
  if sum('YANG Path:' in str(x) or 'CLI Command:' in str(x) for x in r.values()):break
  records.append(rec)
 return records,warns

def find_target_tables(section, prm):
 """Locate actual target-value grids, not the parameter-description matrix.

 CIQ layouts can duplicate target grids (e.g. B:I and K:P); choose the
 leftmost usable copy and retain its sparse column positions.
 """
 candidates=[]
 for ix,(rn,row) in enumerate(section):
  if rn<=prm:continue
  headers=[(c,clean_name(v)) for c,v in sorted(row.items())
           if is_parameter_name(v) and clean_name(v)]
  if len(headers)<2:continue
  # Exclude description rows: parameter name in B, type in C, narrative in G.
  # Require multiple parameter headings on the SAME row.
  preceding=[(j,r) for j,r in section if rn-4<=j<rn]
  evidence=any(any('target values' in str(v).lower() or 'target value (' in str(v).lower()
                   for v in r.values()) for j,r in preceding)
  # Some CIQ tables have broken #REF! in their Target Values heading;
  # the same-row header signature still identifies those tables.
  if not evidence and len(headers)<3:continue
  # Split repeated header grids into groups using original name repetition.
  groups=[];current=[];used=set()
  for col,key in headers:
   if key in used:
    if len(current)>=2:groups.append(current)
    current=[];used=set()
   current.append((col,key));used.add(key)
  if len(current)>=2:groups.append(current)
  for group in groups:
   # Data on immediately following rows; one selected table per section.
   records=[]
   header_cols={col for col,key in group}
   for rr,values in section[ix+1:]:
    if rr>rn+100:break
    if any('YANG Path:' in str(x) or 'CLI Command:' in str(x) for x in values.values()):break
    if any('PARAMATERS' in str(x).upper() or 'PARAMETERS' in str(x).upper() for x in values.values()):break
    if any('Target Values' in str(x) for x in values.values()):break
    if len([v for v in values.values() if is_parameter_name(v)])>=2:break
    rec={key:normalize(values.get(col)) for col,key in group if normalize(values.get(col)) is not None}
    if rec:records.append(rec)
    elif records:break
   if records:candidates.append((rn,group,records,evidence))
 # Prefer evidenced tables, then leftmost duplicate, then most complete table.
 candidates.sort(key=lambda x:(not x[3],x[0],x[1][0][0],-len(x[1]),-len(x[2])))
 return candidates


# Use one eligibility rule across the legacy parser, target-table extractor,
# cross-sheet resolver, and final catalog. Review/Config Status belongs to
# the section immediately before its YANG Path row.
OUT_OF_SCOPE = re.compile(r"\bout[\s_-]*of[\s_-]*scope\b", re.I)
STATUS_LABEL = re.compile(r"\b(?:config|review)\s*status\b", re.I)

def section_is_out_of_scope(rows, path_row):
 """True only when an explicit section status is Out of Scope.

 Inspect the YANG header and preceding status rows; do not scan descriptions
 or nearby unrelated section data. Accept merged cells / shifted columns.
 """
 nearby=[(rn,row) for rn,row in rows if path_row-5 <= rn <= path_row]
 for rn,row in nearby:
  values=[str(v).strip() for _,v in sorted(row.items()) if str(v).strip()]
  for col,value in row.items():
   if STATUS_LABEL.search(str(value)):
    # Status can be in another column of the same row; a textual marker
    # inside a description must not trigger without the status label.
    if any(OUT_OF_SCOPE.search(v) for v in values):return True
 return False

def convert_sheet_legacy(name,rows):
 """Detect full CIQ sections and compact target tables, using clean I:L / I+ target data."""
 data=[]; warnings=[]
 if not rows:return data,warnings
 indexed={i:r for i,r in rows}
 maxrow=rows[-1][0]
 paths=[(i,next((v.split('YANG Path:',1)[1].strip() for v in r.values() if 'YANG Path:' in v),'')) for i,r in rows if any('YANG Path:' in v for v in r.values())]
 if not paths:
  # Compact target-value worksheets: use first header row; choose B+ when headers in B, prefer I+ when duplicate standardized table.
  headers=[(i,r) for i,r in rows[:25] if is_header(r)]
  if not headers:return data,['No parameter header detected']
  hi,head=headers[0]
  cols=[c for c,v in head.items() if clean_name(v) and is_parameter_name(v)]
  if len([c for c in cols if c>=9])>=2:cols=[c for c in cols if c>=9]
  else:cols=[c for c in cols if c>=2]
  records,_=take_table(rows,hi,maxrow+1,head,set(cols))
  if records:data.append((clean_name(name),records,'direct table',hi))
  else:warnings.append('No valid records in compact table')
  return data,warnings
 for ix,(pstart,path) in enumerate(paths):
  pend=paths[ix+1][0] if ix+1<len(paths) else maxrow+1
  # Section status is usually one or two rows immediately before YANG path.
  if section_is_out_of_scope(rows,pstart):continue
  section=[(i,r) for i,r in rows if pstart<i<pend]
  param_rows=[(i,r) for i,r in section if any('PARAMATERS' in str(x).upper() or 'PARAMETERS' in str(x).upper() for x in r.values())]
  if not param_rows:continue
  prm=param_rows[0][0]
  headers=[(i,r) for i,r in section if i>prm and is_header(r)]
  if headers:
   hi,head=headers[0]
   good=[c for c,v in head.items() if c>=9 and is_parameter_name(v) and clean_name(v)]
   if len(good)<2:good=[c for c,v in head.items() if c>=2 and is_parameter_name(v) and clean_name(v)]
   rec,_=take_table(rows,hi,pend,head,set(good))
  else:
   # Vertical target values table, parameter names in B and targets in I.
   rec={}
   for i,r in section:
    if i<=prm or i>prm+75:continue
    k=r.get(2,'')
    if is_parameter_name(k) and not any(word in k.lower() for word in ('yang path','cli command')):
     v=normalize(r.get(9))
     if v is not None:rec[clean_name(k)]=v
   rec=[rec] if rec else []
  if rec:
   leaf=clean_name(path.rstrip('/').split('/')[-1]) or 'config'
   data.append((leaf,rec,path,pstart))
  else:warnings.append(f'{path}: no target values found')
 return data,warnings

def convert_sheet_sparse(name,rows):
 """Detect full CIQ sections and sparse target tables without hardcoded values."""
 data=[];warnings=[]
 if not rows:return data,warnings
 indexed={i:r for i,r in rows}
 maxrow=rows[-1][0]
 paths=[(i,next((v.split('YANG Path:',1)[1].strip() for v in r.values() if 'YANG Path:' in v),'')) for i,r in rows if any('YANG Path:' in v for v in r.values())]
 if not paths:
  headers=[(i,r) for i,r in rows[:25] if is_header(r)]
  if not headers:return data,['No parameter header detected']
  hi,head=headers[0]
  cols=[c for c,v in head.items() if clean_name(v) and is_parameter_name(v)]
  if len([c for c in cols if c>=9])>=2:cols=[c for c in cols if c>=9]
  else:cols=[c for c in cols if c>=2]
  records,_=take_table(rows,hi,maxrow+1,head,set(cols))
  if records:data.append((clean_name(name),records,'direct table',hi))
  else:warnings.append('No valid records in compact table')
  return data,warnings
 for ix,(pstart,path) in enumerate(paths):
  pend=paths[ix+1][0] if ix+1<len(paths) else maxrow+1
  if section_is_out_of_scope(rows,pstart):continue
  section=[(i,r) for i,r in rows if pstart<i<pend]
  param_rows=[(i,r) for i,r in section if any('PARAMATERS' in str(x).upper() or 'PARAMETERS' in str(x).upper() for x in r.values())]
  if not param_rows:continue
  prm=param_rows[0][0]
  tables=find_target_tables(section,prm)
  if tables:
   _,columns,rec,_=tables[0]
  else:
   # Fallback for layouts with target values directly alongside leaf definitions.
   # Do not mistake "Refer to column ..." pointers for literal values.
   vals={}
   for i,r in section:
    if i<=prm or i>prm+75:continue
    k=r.get(2,'')
    if is_parameter_name(k):
     raw=r.get(9)
     if isinstance(raw,str) and raw.lower().startswith('refer to column'):continue
     v=normalize(raw)
     if v is not None:vals[clean_name(k)]=v
   rec=[vals] if vals else []
  if rec:
   leaf=clean_name(path.rstrip('/').split('/')[-1]) or 'config'
   data.append((leaf,rec,path,pstart))
  else:warnings.append(f'{path}: no target values found')
 return data,warnings

def is_supplementary_details_row(row):
 """Detect narrative reference tables after Target Values, irrespective of NF name.

 E.g. 'MME NAS Timer Details' / 'Details hidden' followed by
 'Timers (nas-timer-mme-tid)' / 'Description'.
 """
 cells=[str(v).strip() for _,v in sorted(row.items()) if v is not None and str(v).strip()]
 if not cells:return False
 first=re.sub(r"\s+"," ",cells[0]).strip()
 second=re.sub(r"\s+"," ",cells[1]).strip() if len(cells)>1 else ""
 if re.search(r"(?i)\bdetails\s+hidden\b", first+" "+second):return True
 if re.search(r"(?i)\bdetails\s*$",first) and len(first.split())>=2:return True
 if re.match(r"(?i)^(?:timers?|parameters?|entries|items)\s*\([^)]*\)\s*$",first) and (
     re.match(r"(?i)^(?:description|details|value\s+range|type)\b",second) or len(cells)>=3):return True
 return False


def is_post_target_note(row):
 """Reject explanatory notes following target data, not actual leaf values.

 Notes can appear without blank rows or in the target key column, e.g.
 '* Note: Ensure AMFAUTH-CONF is aligned with these values'.
 Only inspect the leading nonempty text; avoid excluding string values
 containing the word 'note' elsewhere in a legitimate field.
 """
 cells=[str(v).replace('\u00a0',' ').strip() for _,v in sorted(row.items())
        if v is not None and str(v).strip()]
 if not cells:return False
 first=re.sub(r'\s+', ' ', cells[0]).strip()
 # Excel copy/pastes may use literal stars, bullets, or numbered NOTE: labels.
 return bool(re.match(r'(?i)^(?:[\*•▪‣-]\s*)?(?:note\s*[:：]|notes\s*[:：]|remarks?\s*[:：]|\*\s*note\b)',first))


def find_vertical_target_values(section, prm):
 """Read vertical target blocks, keeping *every* repeated child-list instance.

 Child-list instances begin when a declared child key leaf recurs (or, if
 the CIQ does not identify one, when the first child field recurs). This
 handles SGW/PGW/SCP groups, EIR action lists and other vertical tables.
 Parent scalar values are carried to each instance; [ENC] is skipped.
 """
 declared=declared_section_leaves(section,prm)
 known=set(declared)
 child_keys=[]
 for rr,row in section:
  if rr<=prm:continue
  if any(re.search(r"\bTarget Values\s*\(",str(v),re.I) for v in row.values()):break
  raw_key=re.sub(r"\s*/\s*",'/',str(row.get(2,'')).strip())
  key=clean_name(raw_key)
  typ=re.sub(r'\s+',' ',str(row.get(3,''))).lower()
  if '/' in raw_key and key in known and 'key leaf' in typ:
   child_keys.append(key)
 first_child=next((k for k in declared if k in known and
                   any(k==clean_name(re.sub(r"\s*/\s*",'/',str(r.get(2,''))))
                       and '/' in str(r.get(2,'')) for _,r in section if _>prm)),None)
 tables=[]
 for pos,(rn,row) in enumerate(section):
  headings=[c for c,v in sorted(row.items()) if re.search(r"\bTarget Values\s*\(",str(v),re.I)]
  if rn<=prm or not headings:continue
  keycol=headings[0]
  currentcols=[c for c,v in row.items() if re.search(r"\bCurrent Values\s*\(",str(v),re.I)]
  boundary=min((c for c in currentcols if c>keycol),default=9999)
  if keycol+1>=boundary:continue
  # Child keys are scoped to their declared YANG depth. A repeated inner
  # child key starts a new inner instance without losing its ancestor service.
  # A repeated outer key starts an entirely new outer instance.
  depth_by_leaf={}
  first_at_depth={}
  for declared_key in declared:
   if declared_key in known:
    # A normalized full child path has one underscore-joined token per
    # original slash segment. Read the actual declaration to get its depth.
    matching=next((str(r.get(2,'')) for rnum,r in section if prm<rnum<rn
                   and clean_name(re.sub(r'\s*/\s*','/',str(r.get(2,''))))==declared_key),None)
    if matching and '/' in matching:
     depth=matching.count('/')+1
     depth_by_leaf[declared_key]=depth
     first_at_depth.setdefault(depth,declared_key)
  parent={};child={};records=[]
  for rr,values in section[pos+1:]:
   if rr>rn+500:break
   if is_supplementary_details_row(values) or is_post_target_note(values):break
   if any(re.search(r"\b(?:Target Values\s*\(|PARAMATERS|PARAMETERS|YANG\s*Path\s*:)",str(v),re.I) for v in values.values()):break
   raw_key=re.sub(r"\s*/\s*",'/',str(values.get(keycol,'')).strip())
   key=clean_name(raw_key)
   if key not in known:continue
   val=normalize(values.get(keycol+1))
   if val is None:continue
   if isinstance(val,str) and re.match(r'^(refer to|see value below)',val,re.I):continue
   if '/' in raw_key:
    depth=depth_by_leaf.get(key,raw_key.count('/')+1)
    is_key=(key in child_keys or (not child_keys and key==first_at_depth.get(depth)))
    # Repeating a declared key starts a new instance at that hierarchy level.
    # Keep shallower ancestor values (e.g. service name/scheme) whenever a
    # deeper list instance (e.g. service IP) starts. Do NOT keep its sibling
    # values when a new outer list instance starts.
    if child and key in child and (is_key or child[key]!=val):
     records.append({**parent,**child})
     child={k:v for k,v in child.items() if depth_by_leaf.get(k,2)<depth}
    child[key]=val
   else:
    # A new parent starts a new record. Never carry child objects from
    # one parent into another or replace an earlier completed group.
    if key in parent and parent[key]!=val:
     if parent or child:records.append({**parent,**child})
     parent={};child={}
    parent[key]=val
  if parent or child:records.append({**parent,**child})
  tables.extend((rn,{k:rec[k] for k in declared if k in rec}) for rec in records)
 return tables


def declared_section_leaves(section, prm):
 """Extract declared leaf names from the PARAMATERS matrix, including bare names.

 Only a row with a recognized Type cell may declare a leaf. This avoids
 confusing narrative text, column hints or Target/Current headings with leaves.
 """
 declared=[]
 for rn, row in section:
  if rn<=prm:continue
  if any(re.search(r"\bTarget Values\s*\(",str(v),re.I) for v in row.values()):break
  key=str(row.get(2,'')).strip()
  typ=str(row.get(3,'')).strip().lower()
  if not re.search(r"\b(?:key\s+leaf|leaf|child\s+list|leaf-list)\b",typ):continue
  # Bare YANG names such as pod, prefix, gateway and distance are legal.
  if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_.:-]*(?:\s*/\s*[A-Za-z][A-Za-z0-9_.:-]*)*",key):continue
  norm=clean_name(re.sub(r"\s*/\s*",'/',key))
  if norm and norm not in declared:declared.append(norm)
 return declared

def is_continuation_marker(value):
 """CIQ arrow glyphs mean carry forward the previous parent key/value."""
 if value is None:return False
 token=str(value).strip().replace('\xa0','')
 return bool(token) and bool(re.fullmatch(r'[↪⇢↳→➜⇒↦⤷]+',token))


def repair_shifted_target_cells(values,columns,declared,section,prm):
 """Recover a displaced last text/numeric pair, without changing valid columns.

 Some CIQs have a blank cause cell followed by a text cause in the numeric
 NASC column and a numeric value in the immediate following spare column.
 Use the declared types to ensure the shift is unambiguous.
 """
 result=dict(values)
 if len(columns)<2:return result
 numeric=set()
 text=set()
 for rn,row in section:
  if rn<=prm:continue
  if any('Target Values' in str(v) for v in row.values()):break
  key=clean_name(row.get(2,''))
  if key not in declared:continue
  desc=str(row.get(7,'')).lower()
  if re.search(r'\btype\s*:\s*(?:u?int(?:8|16|32|64)?\b)',desc):numeric.add(key)
  if re.search(r'\btype\s*:\s*(?:string\b|[^\n]*,\s*string\b)',desc):text.add(key)
 for i in range(1,len(columns)):
  prevcol,prevkey=columns[i-1];col,key=columns[i]
  if prevkey not in text or key not in numeric:continue
  if normalize(result.get(prevcol)) is not None:continue
  shifted=normalize(result.get(col)); next_value=normalize(result.get(col+1))
  if shifted is None or next_value is None:continue
  if not re.fullmatch(r'[+-]?\d+',str(next_value)):continue
  if re.fullmatch(r'[+-]?\d+',str(shifted)):continue
  result[prevcol]=result[col];result[col]=result[col+1]
 return result


def find_horizontal_target_values(section, prm):
 """Read multi-row target grids using the declared parameter/leaf schema.

 The Target Values heading defines the left grid; Current Values defines the
 exclusive right boundary. Sparse columns, spacer columns, review symbols and
 human notes are ignored. Every populated row is a separate record.
 """
 declared=declared_section_leaves(section,prm)
 if len(declared)<2:return []
 known=set(declared)
 results=[]
 for idx,(rn,row) in enumerate(section):
  target_starts=[c for c,v in sorted(row.items()) if re.search(r"\bTarget Values\s*\(",str(v),re.I)]
  if rn<=prm or not target_starts:continue
  target_start=min(target_starts)
  current_starts=[c for c,v in sorted(row.items()) if re.search(r"\bCurrent Values\s*\(",str(v),re.I)]
  current_start=min((c for c in current_starts if c>target_start),default=10000)
  header_at=None; columns=[]
  for j in range(idx+1,min(idx+6,len(section))):
   hr,head=section[j]
   if hr>rn+5:break
   candidates=[];seen=set()
   for col,value in sorted(head.items()):
    if not target_start<=col<current_start:continue
    key=clean_name(re.sub(r"\s*/\s*",'/',str(value).strip()))
    if key in known and key not in seen:
     candidates.append((col,key));seen.add(key)
   if len(candidates)>=2:
    header_at=j; columns=candidates;break
  if header_at is None:continue
  records=[];blank=0
  previous_values={}
  colset={c for c,_ in columns}
  previous_row=section[header_at][0]
  for rr,values in section[header_at+1:]:
   if rr>section[header_at][0]+5000:break
   # Empty Excel rows are not included in the sparse row iterator.  A gap
   # *after* valid target records terminates this target table.  Otherwise
   # the following explanatory/reference table can be mistaken for data.
   if records and rr>previous_row+1:break
   previous_row=rr
   if any(re.search(r"\b(?:Target Values\s*\(|Current Values\s*\(|PARAMATERS|PARAMETERS|YANG\s*Path\s*:)",str(v),re.I) for v in values.values()):break
   if sum(clean_name(values.get(c,'')) in known for c in colset)>=2:break
   # Stop at the first supplementary details/reference table.
   if is_supplementary_details_row(values) or is_post_target_note(values):break
   record={}
   corrected=repair_shifted_target_cells(values,columns,declared,section,prm)
   for col,key in columns:
    raw=corrected.get(col)
    val=previous_values.get(key) if is_continuation_marker(raw) else normalize(raw)
    if val is not None:record[key]=val
   for key,value in record.items():previous_values[key]=value
   # A valid record needs its first/key leaf; avoids standalone notes.
   if not record or columns[0][1] not in record:
    blank+=1
    if records and blank>=1:break
    continue
   records.append({key:record[key] for key in declared if key in record})
   blank=0
  if records:results.append((rn,records))
 return results

def convert_sheet(name,rows):
 """Prefer verified target blocks over fallback/description-matrix placeholders."""
 # For standalone sheets, reject a duplicated Current Values grid and extract
 # ALL records from the leftmost set of header columns (not merely the rows
 # where both left and right tables happen to be populated).
 if rows and not any(any('YANG Path:' in str(v) for v in row.values()) for _,row in rows):
  for rn,header in rows[:35]:
   seen=set();columns=[];has_duplicate=False
   for col,value in sorted(header.items()):
    key=clean_name(re.sub(r'\s*/\s*','/',str(value).strip()))
    if key in seen and len(columns)>=2:
     has_duplicate=True;break
    if is_parameter_name(value) and col>=2:
     columns.append(key);seen.add(key)
   if has_duplicate and len(columns)>=2:
    notes=[]
    records=extract_reference_grid(rows,columns,notes,name)
    if records:
     return [(clean_name(name),records,'direct table',rn)], notes
 baseline,base_warnings=convert_sheet_legacy(name,rows)
 upgraded,_=convert_sheet_sparse(name,rows)
 old_map={(path,rn):i for i,(_,_,path,rn) in enumerate(baseline)}
 for leaf,records,path,rn in upgraded:
  ident=(path,rn)
  if ident not in old_map:continue
  old_records=baseline[old_map[ident]][1]
  if len(old_records)==len(records) and all(set(a)==set(b) for a,b in zip(old_records,records)):
   baseline[old_map[ident]]=(leaf,records,path,rn)
 # A verified target-value table is authoritative even when the legacy
 # extraction found the same number of spurious `See value below.` fields.
 paths=[(i,next((str(v).split('YANG Path:',1)[1].strip() for v in r.values() if 'YANG Path:' in str(v)),''))
        for i,r in rows if any('YANG Path:' in str(v) for v in r.values())]
 maxrow=rows[-1][0]+1 if rows else 1
 replacements={};new=[]
 for idx,(start,path) in enumerate(paths):
  if section_is_out_of_scope(rows,start):continue
  end=paths[idx+1][0] if idx+1<len(paths) else maxrow
  section=[(i,r) for i,r in rows if start<i<end]
  prm=next((i for i,r in section if any(re.fullmatch(r'PARAMATERS|PARAMETERS',str(v).strip(),re.I) for v in r.values())),None)
  if prm is None:continue
  horizontal=find_horizontal_target_values(section,prm)
  if horizontal:
   # Each row is one record. Do not consume duplicated Current Values columns.
   recs=horizontal[0][1]
  else:
   tables=find_vertical_target_values(section,prm)
   if not tables:continue
   recs=[record for _,record in tables]
  # Prefer the first populated target block. Later blocks are additional
  # records only when they also contain real assigned values.
  ident=(path,start)
  replacement=(clean_name(path.rstrip('/').split('/')[-1]),recs,path,start)
  if ident in old_map:replacements[old_map[ident]]=replacement
  else:new.append(replacement)
 for i,replacement in replacements.items():baseline[i]=replacement
 baseline.extend(new)
 return baseline,base_warnings

def discover_key_value_table(name,rows):
 """Support new simple worksheets with Parameter/Target Value columns."""
 for ix,(rn,row) in enumerate(rows[:100]):
  labels={c:str(v).lower().strip().replace(' ','_') for c,v in row.items()}
  keycols=[c for c,v in labels.items() if v in ('parameter','parameter_name','config_parameter','yang_parameter','key')]
  valcols=[c for c,v in labels.items() if v in ('target_value','target_values','value','config_value','desired_value')]
  if not keycols or not valcols:continue
  keycol,valcol=keycols[0],valcols[0]
  entries=[]
  for rr,values in rows[ix+1:]:
   k=values.get(keycol,'').strip()
   if not k:continue
   if k.lower() in ('parameter','parameter_name'):break
   value=normalize(values.get(valcol))
   if value is None or not clean_name(k):continue
   entries.append({clean_name(k):value})
  if entries:
   # One map assembled from the parameter column; repeated keys produce records.
   result={}; multiples=[]
   for entry in entries:
    k,v=next(iter(entry.items()))
    if k in result:
     if result: multiples.append(result);result={}
    result[k]=v
   if result:multiples.append(result)
   return [(clean_name(name),multiples,'parameter/value table',rn)],[]
 return [],['No recognizable YANG blocks, parameter/value pairs, or compact tables']

def collect_section_metadata(rows):
 """Read visible CIQ section metadata and declared leaf order without fixed leaf names."""
 indexed={rn:row for rn,row in rows}
 paths=[]
 for rn,row in rows:
  for value in row.values():
   m=re.search(r"YANG\s*Path\s*:\s*(.*)",str(value),re.I)
   if m:
    paths.append((rn,m.group(1).strip()));break
 result={}
 for i,(start,path) in enumerate(paths):
  end=paths[i+1][0] if i+1<len(paths) else (rows[-1][0]+1 if rows else start+1)
  section=[(rn,row) for rn,row in rows if start<=rn<end]
  metadata={'yang_path':path,'cli_command':'','description':'','dependencies':'','parameter_order':[],'child_paths':{},'string_leaves':[],'key_leaves':[],'numeric_leaves':[]}
  state=None
  for rn,row in section:
   values=[str(v).strip() for _,v in sorted(row.items()) if str(v).strip()]
   if not values:continue
   first=values[0]
   cli=re.search(r'CLI\s*Command\s*:\s*(.*)',first,re.I)
   if cli:metadata['cli_command']=cli.group(1).strip();state=None;continue
   if re.match(r'^DESCRIPTION\s*$',first,re.I):state='description';continue
   if re.match(r'^DEPENDENCIES\s*$',first,re.I):state='dependencies';continue
   if re.match(r'^PARAMATERS\s*$|^PARAMETERS\s*$',first,re.I):state='parameters';continue
   if state in ('description','dependencies'):
    if re.match(r'^(Environment Specific\?|Config Status|Target Values|Review Status)',first,re.I):state=None;continue
    # Descriptive content occurs predominantly in column B (or first populated cell).
    metadata[state]+=('\n' if metadata[state] else '')+first
   elif state=='parameters':
    leaf=row.get(2,'')
    if (is_parameter_name(leaf) or (re.fullmatch(r'[A-Za-z][A-Za-z0-9_.:-]*',str(leaf).strip()) and re.search(r'\b(?:key\s+leaf|leaf|leaf-list|child\s+list)\b',str(row.get(3,'')).lower()))):
     key=clean_name(leaf)
     if key not in metadata['parameter_order']:metadata['parameter_order'].append(key)
     if 'key leaf' in str(row.get(3,'')).lower() and key not in metadata['key_leaves']:metadata['key_leaves'].append(key)
     definition=str(row.get(7,'')).lower().replace('_x000d_', '\n')
     if re.search(r'\btype\s*:\s*(?:string\b|[^\n]*,\s*string\b)',definition) and key not in metadata['string_leaves']:
      metadata['string_leaves'].append(key)
     if re.search(r'\btype\s*:\s*(?:u?int(?:8|16|32|64)?\b)',definition) and key not in metadata['numeric_leaves']:
      metadata['numeric_leaves'].append(key)
     if '/' in str(leaf):
      parts=[clean_name(part.strip()) for part in str(leaf).split('/') if part.strip()]
      if len(parts)>=2:metadata['child_paths'][key]=parts
    elif re.match(r'^Target Values',first,re.I):state=None
  result[(path,start)]=metadata
 return result

def reorder_entries(entries,leaf_order):
 """Ensure defined YANG leaf order is the same for all records, then retain extra input keys."""
 ordered=[]
 for entry in entries:
  ordered.append({key:entry[key] for key in leaf_order if key in entry} |
                 {key:value for key,value in entry.items() if key not in leaf_order})
 return ordered

# Cross-sheet reference resolution is intentionally schema-driven; no PLMN-specific
# value lists or leaf names are hardcoded here.
# Detect CIQ cross-worksheet pointers at the section heading, target field,
# or explicit trailing instruction. Not tied to any network element or tab.
SHEET_REFERENCE = re.compile(
 r"(?:^|\b)(?:[=\-]\s*>\s*)?refer\s+to\s+(.+?)\s+sheet\s*\.?\s*$",
 re.I,
)

def referenced_sheet(value):
 if not isinstance(value,str):return None
 value=re.sub(r"\s+", " ", value.replace("\xa0", " ")).strip()
 match=SHEET_REFERENCE.search(value)
 if not match:return None
 target=match.group(1).strip(" .\t\"'\u2018\u2019\u201c\u201d")
 return target or None

def _sheet_token(value):
 # Normalize spaces, case, underscores and punctuation, including in names
 # broken by spreadsheet formatting (e.g. 'PGW  Groups').
 return re.sub(r"[^a-z0-9]+", "", str(value).casefold())

def _match_referenced_worksheet(target, sheet_rows):
 """Find an unambiguous worksheet by its complete normalized name.

 Never use approximate substring/fuzzy matching: a wrong worksheet can
 generate deployable configuration with unintended values.
 """
 token=_sheet_token(target)
 hits=[(name, rows) for name,rows in sheet_rows.items() if _sheet_token(name)==token]
 if len(hits)==1:return hits[0],None
 if len(hits)>1:return None,f"ambiguous worksheet name '{target}' ({', '.join(n for n,_ in hits)})"
 return None,f"referenced worksheet '{target}' not found"

def section_level_sheet_references(rows):
 """Find a CIQ section's external target worksheet without reading other sections.

 References can be on a Review Status heading or after the leaf definitions.
 The next Review Status row closes the current section even before the next
 YANG Path, preventing a later section's pointer from being misassigned.
 """
 paths=[rn for rn,row in rows if any(re.search(r"YANG\s*Path\s*:",str(v),re.I) for v in row.values())]
 found={}
 for i,start in enumerate(paths):
  prev=paths[i-1] if i else 0
  next_path=paths[i+1] if i+1<len(paths) else (rows[-1][0]+1 if rows else start+1)
  next_header=next((rn for rn,row in rows if start<rn<next_path and
    any(re.search(r"Review\s*Status",str(v),re.I) for v in row.values())),next_path)
  # Own preamble starts at nearest Review Status preceding this YANG Path.
  header=next((rn for rn,row in reversed(rows) if prev<rn<start and
    any(re.search(r"Review\s*Status",str(v),re.I) for v in row.values())),start)
  candidates=[]
  for rn,row in rows:
   if header<=rn<next_header:
    for value in row.values():
     target=referenced_sheet(value)
     if target:candidates.append(target)
  unique={_sheet_token(t):t for t in candidates}
  if len(unique)==1:found[start]=next(iter(unique.values()))
  elif len(unique)>1:found[start]=list(unique.values())
 return found

def _reference_leaf_alias(header, declared):
 """Resolve an exact leaf or a uniquely shortened YANG leaf path.

 Some value-only CIQ worksheets explicitly remove a shared namespace prefix
 (e.g. amfmme-ne-grp -> ne-grp). Never guess when two declarations share a
 shortened suffix; missing or ambiguous references remain unresolved.
 """
 key=clean_name(re.sub(r"\s*/\s*", "/", str(header).strip()))
 if key in declared:return key
 candidates=[]
 namespace_prefixes={full.split('_')[0]+'_' for full in declared if '_' in full}
 for full in declared:
  if (key and full.endswith('_'+key)) or any(full.replace(prefix,'')==key for prefix in namespace_prefixes):
   candidates.append(full)
 return candidates[0] if len(candidates)==1 else None


def extract_reference_grid(rows, wanted_keys, warnings, sheet_name):
 """Parse only target-side data, honoring shortened headers and CIQ arrows.

 Retain one physical record per populated row; repeated parent values are
 carried forward explicitly and grouped by YANG list keys at render time.
 """
 known=list(dict.fromkeys(wanted_keys))
 options=[]
 for ix,(rn,row) in enumerate(rows):
  header=[];used=set();aliases=[]
  for col,value in sorted(row.items()):
   key=_reference_leaf_alias(value,known)
   if key is None:
    from difflib import get_close_matches
    cleaned=clean_name(value)
    close=get_close_matches(cleaned,known,n=2,cutoff=0.9)
    if len(close)==1 and abs(len(close[0])-len(cleaned))<=1 and cleaned.split('_')[0]==close[0].split('_')[0]:
     key=close[0]
     warnings.append(f'{sheet_name}: mapped header {cleaned} to declared {key}')
   if key is not None:
    if key in used:break # Duplicate headers introduce Current Values grid.
    header.append((col,key));used.add(key)
    if clean_name(value)!=key:aliases.append((value,key))
  if len(header)<2 or known[0] not in used:continue
  records=[];previous={};last_rr=rn
  for rr,vals in rows[ix+1:]:
   if rr>rn+100000:break
   if records and rr>last_rr+1:break
   last_rr=rr
   if is_supplementary_details_row(vals) or is_post_target_note(vals):break
   if any(re.search(r'\b(?:Target Values\s*\(|PARAMATERS|PARAMETERS|YANG\s*Path\s*:)',str(v),re.I) for v in vals.values()):break
   rec={}
   for col,key in header:
    raw=vals.get(col)
    if is_continuation_marker(raw):
     value=previous.get(key)
    else:
     value=normalize(raw)
    if value is not None:rec[key]=value
   if known[0] in rec and len(rec)>=2:
    records.append({k:rec[k] for k in known if k in rec})
    previous.update(rec)
   elif records:
    # Blank/marker-only row or secondary notes mark table end.
    break
  if records:options.append((len(used),len(records),-header[0][0],records,aliases))
 if not options:return []
 winner=max(options,key=lambda x:x[:3])
 # Diagnostic only; avoid repeated warnings for every row.
 return winner[3]

def resolve_sheet_references(catalog, sheet_rows, warnings):
 """Follow any declared `Refer to <sheet> sheet` pointer generically.

 A full CIQ section can mention its target worksheet in the header, in
 a target-value placeholder, or in a footer. Keep the originating YANG
 definition as the source of names/types/nested lists. No placeholder
 text may become a Terraform assignment.
 """
 for block in catalog:
  entries=block['entries']
  pointers={referenced_sheet(v) for record in entries for v in record.values() if referenced_sheet(v)}
  explicit=block.get('declared_reference')
  if explicit:pointers.add(explicit)
  if not pointers:continue
  unique={_sheet_token(p):p for p in pointers}
  if len(unique)!=1:
   warnings.append(f"{block['sheet']} / {block['config']}: ambiguous referenced sheets: {sorted(pointers)}")
   block['entries']=[{k:v for k,v in r.items() if referenced_sheet(v) is None} for r in entries]
   continue
  target=next(iter(unique.values()))
  hit,error=_match_referenced_worksheet(target,sheet_rows)
  if error:
   warnings.append(f"{block['sheet']} / {block['config']}: {error}")
   block['entries']=[{k:v for k,v in r.items() if referenced_sheet(v) is None} for r in entries]
   continue
  target_name,raw_rows=hit
  declared=block.get('metadata',{}).get('parameter_order') or list(dict.fromkeys(k for r in entries for k in r))
  grid=extract_reference_grid(raw_rows,declared,warnings,target_name)
  if not grid:
   warnings.append(f"{block['sheet']} / {block['config']}: no matching target table found in '{target_name}'")
   block['entries']=[{k:v for k,v in r.items() if referenced_sheet(v) is None} for r in entries]
   continue
  block['entries']=grid
  block['reference_sheet']=target_name
 return catalog

def enrich_referenced_direct_tables(catalog, warnings):
 """Inherit YANG declarations for directly selected referenced worksheets.

 Prefer the already validated cross-sheet result to values-only direct tables,
 so shortened target headers, continuation arrows and nested lists all retain
 their correct interpretation in independent sheet-selection mode as well.
 """
 by_sheet={}
 for section in catalog:
  target=section.get('reference_sheet')
  if target and section.get('metadata',{}).get('parameter_order'):
   by_sheet.setdefault(_sheet_token(target),[]).append(section)
 for direct in catalog:
  if direct.get('path')!='direct table':continue
  matches=by_sheet.get(_sheet_token(direct['sheet']),[])
  if len(matches)!=1:
   if len(matches)>1:warnings.append(f"{direct['sheet']}: multiple referring sections; cannot infer unique YANG schema")
   continue
  canonical=matches[0]
  direct['config']=canonical['config']
  direct['path']=canonical['path']
  direct['metadata']=dict(canonical['metadata'])
  direct['entries']=reorder_entries(canonical['entries'],direct['metadata']['parameter_order'])
  direct['inherited_schema_from']=canonical['sheet']
 return catalog

def discover(workbook):
 catalog=[]; errors=[]; sheets=[]; sheet_rows={}
 for sheet,rows in read_workbook(workbook):
  sheet_rows[sheet]=rows
  sheets.append(sheet)
  blocks,warnings=convert_sheet(sheet,rows)
  if not blocks:
   blocks,extra=discover_key_value_table(sheet,rows)
   warnings=extra if blocks else warnings+extra
  metadata=collect_section_metadata(rows)
  section_refs=section_level_sheet_references(rows)
  # Sections with only "See value below" in their leaf definition table have
  # no local target records. Represent them as pending cross-sheet resolutions.
  found={(path,rn) for _,_,path,rn in blocks}
  for (path,rn),meta in metadata.items():
   if (path,rn) in found or section_is_out_of_scope(rows,rn):continue
   ref=section_refs.get(rn)
   if not isinstance(ref,str):
    if isinstance(ref,list):warnings.append(f"{path}: ambiguous sheet pointers: {ref}")
    continue
   order=meta.get('parameter_order',[])
   if not order:
    warnings.append(f"{path}: no declared leaves for referenced worksheet '{ref}'")
    continue
   blocks.append((clean_name(path.rstrip('/').split('/')[-1]),[{order[0]:f'Refer to {ref} sheet'}],path,rn))
  for leaf,entries,path,rn in blocks:
   # Defense in depth: never put a skipped CIQ section in the catalog.
   if path != 'direct table' and section_is_out_of_scope(rows,rn):
    continue
   meta=metadata.get((path,rn),{})
   entries=reorder_entries(entries,meta.get('parameter_order',[]))
   block={'sheet':sheet,'config':clean_name(leaf),'entries':entries,'path':path,'row':rn,'metadata':meta}
   # A section-level pointer may be present even when the legacy extractor
   # returns a nonempty (but placeholder-only) entry. Keep it explicitly;
   # relying on extracted values alone misses such cross-sheet sections.
   if path != 'direct table':
    ref=section_refs.get(rn)
    if isinstance(ref,str):block['declared_reference']=ref
    elif isinstance(ref,list):
     warnings.append(f'{path}: ambiguous sheet pointers: {ref}')
   catalog.append(block)
  errors.extend(f'{sheet}: {w}' for w in warnings)
 resolve_sheet_references(catalog,sheet_rows,errors)
 enrich_referenced_direct_tables(catalog,errors)
 catalog=[block for block in catalog if block["entries"] and any(block["entries"])]
 return sheets,catalog,errors

def matches(value,patterns):
 return any(fnmatch.fnmatchcase(value.casefold(),p.casefold()) for p in patterns)

def csv_patterns(value):
 return [x.strip() for x in (value or '').split(',') if x.strip()]

def choose_indices(prompt,items):
 print('\n'+prompt)
 print('  0. All')
 for i,v in enumerate(items,1):print(f'  {i}. {v}')
 while True:
  answer=input('Select 0, numbers (e.g. 1,3-5), or names/patterns: ').strip()
  if not answer or answer.lower() in ('0','all','*'):return None
  selected=set();invalid=[]
  for term in csv_patterns(answer):
   if re.fullmatch(r'\d+-\d+',term):
    a,b=map(int,term.split('-'))
    if not (1<=a<=b<=len(items)):invalid.append(term)
    else:selected.update(items[a-1:b])
   elif term.isdigit():
    num=int(term)
    if 1<=num<=len(items):selected.add(items[num-1])
    else:invalid.append(term)
   else:
    found=[item for item in items if matches(item,[term])]
    if found:selected.update(found)
    else:invalid.append(term)
  if not invalid and selected:return selected
  print('Unrecognized selection:',', '.join(invalid) or answer)

def _comment_lines(value):
 """Preserve worksheet prose as ### comment lines without injecting HCL syntax."""
 lines=[]
 for raw in str(value).replace('\r\n','\n').replace('\r','\n').split('\n'):
  line=raw.strip()
  if line:
   # Existing hashes in the source should not change the template prefix.
   lines.append('### '+line)
 return lines

def render_multilevel_group(variable,records,source,metadata):
 """Render repeated CIQ rows as arbitrarily deep YANG child-list hierarchies.

 Root and child lists are grouped by their declared key leaves. They are not
 grouped by values from arbitrary non-key fields. All levels preserve input
 order; repeated service names accumulate individual cause entries.
 """
 from collections import OrderedDict
 child_paths=metadata['child_paths']
 key_fields=set(metadata.get('key_leaves',[]))
 string_fields=set(metadata.get('string_leaves',[]))
 path_map={key:parts for key,parts in child_paths.items()}
 for key in metadata.get('parameter_order',[]):path_map.setdefault(key,[key])
 roots=OrderedDict()

 def identify_child(row,prefix,declared_keys,options):
  """Identify a child without dropping optional/unset YANG key fields.

  Prefer *complete* declared keys.  When a key was omitted (e.g. [ENC]
  or an unconfigured API prefix), use the other direct leaf values only as
  a provisional identity. Never invent a missing leaf or silently overwrite.
  """
  actual=[k for k in declared_keys if k in row and row[k] is not None]
  if len(actual)==len(declared_keys) and actual:
   return tuple((k,row[k]) for k in actual)
  siblings=[k for k in metadata.get('parameter_order',[])
            if k in row and tuple(path_map[k][:-1])==prefix and k not in declared_keys]
  available=actual+siblings
  if not available:
   # No direct identity: only attach if an ancestor uniquely determines this
   # child. More than one possible child is unsafe and remains an error.
   if len(options)==1 and not declared_keys:
    return options[0]['identity']
   missing=[k for k in declared_keys if k not in row]
   raise ValueError(f'{variable}: missing child-list key(s) {missing} for {"/".join(prefix)}; no usable scalar to infer instance')
  identity=(('__inferred_identity__',True),)+tuple((k,row[k]) for k in available)
  msg=(f'{variable}: {"/".join(prefix)} missing declared key(s) '
       f'{[k for k in declared_keys if k not in row]}; using available child values '
       'as provisional identity; validate against YANG before applying')
  if msg not in metadata.setdefault('validation_warnings',[]):
   metadata['validation_warnings'].append(msg)
  return identity

 def add_child(node,path,fullkey,value):
  if len(path)==1:
   leaf=path[0]
   existing=node['values'].get(leaf)
   if existing is not None and existing!=value:
    raise ValueError(f'Conflicting values for {fullkey}: {existing!r} versus {value!r}')
   node['values'][leaf]=(fullkey,value)
   return
  container=path[0]
  candidates=node['children'].setdefault(container,[])
  # Children are selected per target row, not from leaf value alone.
  return container,candidates,path[1:]

 # For each raw row select each level by the key leaf in its declared path.
 for row in records:
  root_keys=[k for k in metadata.get('parameter_order',[]) if k in key_fields and k not in child_paths]
  root_id=tuple((k,row.get(k)) for k in root_keys)
  if not root_id or any(v is None for _,v in root_id):
   raise ValueError(f'{variable}: missing top-level key in target row: {row}')
  root=roots.setdefault(root_id,{'values':OrderedDict(),'children':OrderedDict(),'ids':{}})
  # First collect scalar leaves. Follow metadata order for deterministic output.
  for fullkey in metadata.get('parameter_order',[]):
   if fullkey in row and fullkey not in child_paths:
    item=(fullkey,row[fullkey])
    old=root['values'].get(fullkey)
    if old is not None and old!=item:raise ValueError(f'{variable}: conflicting parent {fullkey}')
    root['values'][fullkey]=item
  containers=OrderedDict()
  for fullkey in metadata.get('parameter_order',[]):
   if fullkey in row and fullkey in child_paths:
    parts=path_map[fullkey]
    for level in range(1,len(parts)):
     prefix=tuple(parts[:level]);containers.setdefault(prefix,[]).append((fullkey,parts))
  for prefix in sorted(containers,key=len):
   parent=root
   for ancestor in range(1,len(prefix)):
    ancestor_path=prefix[:ancestor]
    token=ancestor_path[-1]
    options=parent['children'].setdefault(token,[])
    declared_keys=[k for k in metadata.get('parameter_order',[]) if k in key_fields and tuple(path_map[k][:-1])==ancestor_path]
    identity=identify_child(row,ancestor_path,declared_keys,options)
    child=next((x for x in options if x['identity']==identity),None)
    if child is None:
     child={'values':OrderedDict(),'children':OrderedDict(),'identity':identity}
     options.append(child)
    parent=child
   token=prefix[-1]
   options=parent['children'].setdefault(token,[])
   declared_keys=[k for k in metadata.get('parameter_order',[]) if k in key_fields and tuple(path_map[k][:-1])==prefix]
   if not declared_keys:
    # Older CIQs sometimes mark no list key; use its first declared leaf.
    declared_keys=next(([k] for k in metadata.get('parameter_order',[])
                        if k in row and tuple(path_map[k][:-1])==prefix),[])
   identity=identify_child(row,prefix,declared_keys,options)
   child=next((x for x in options if x['identity']==identity),None)
   # Reject inconsistent rows reusing an explicit child key. With an
   # unkeyed list, differing first-leaf values are separate identities.
   if child is None:
    child={'values':OrderedDict(),'children':OrderedDict(),'identity':identity}
    options.append(child)
   for fullkey,parts in containers[prefix]:
    if tuple(parts[:-1])!=prefix:continue
    leaf=parts[-1];old=child['values'].get(leaf)
    item=(fullkey,row[fullkey])
    if old is not None and old!=item:
     raise ValueError(f'{variable}: child list {token}, identity {identity}: conflicting {fullkey}; no values were overwritten')
    child['values'][leaf]=item

 lines=['####################']
 path=metadata.get('yang_path') or (source.split(' | ')[1] if ' | ' in source else '')
 if path:lines.extend(_comment_lines('Xpath: '+path))
 if metadata.get('cli_command'):lines.extend(_comment_lines('CLI Commands: '+metadata['cli_command']))
 if metadata.get('description'):
  lines.append('### Description:');lines.extend(_comment_lines(metadata['description']))
 if metadata.get('dependencies'):
  dep=str(metadata['dependencies']).splitlines()
  lines.extend(_comment_lines('Dependencies: '+dep[0]))
  for item in dep[1:]:lines.extend(_comment_lines(item))
 lines+=['####################','',f'{variable} = [']
 def render_node(node,indent):
  pad=' '*indent
  lines.append(pad+'{')
  active=[(k,v) for k,(full,v) in node['values'].items() if not isinstance(v,DesignDefault)]
  default=[(k,v) for k,(full,v) in node['values'].items() if isinstance(v,DesignDefault)]
  w=max([len(k) for k in node['values']] or [0])
  for k,v in active:
   full=node['values'][k][0]
   formatted=json.dumps(str(v),ensure_ascii=False) if full in string_fields else hcl(v)
   lines.append(pad+f'  {k.ljust(w)} = {formatted}')
  if default:
   lines.append(pad+'  ### Defaults used by design')
   for k,v in default:
    full=node['values'][k][0]
    formatted=json.dumps(str(v),ensure_ascii=False) if full in string_fields else hcl(v)
    lines.append(pad+f'  # {k.ljust(w)} = {formatted}')
  for container,children in node['children'].items():
   lines.append(pad+f'  {container} = [')
   for child in children:render_node(child,indent+4)
   lines.append(pad+'  ]')
  if indent==2:
   lines.extend([pad+'  terrayang = {',pad+'    # with_defaults = "report-all"',pad+'  }'])
  lines.append(pad+'},')
 for root in roots.values():render_node(root,2)
 lines+= [']','']
 return '\n'.join(lines)


def render_group(variable,records,source,metadata=None):
 """Emit CIQ metadata, ordered scalar leaves and nested child-list entries."""
 metadata=metadata or {}
 # Every declared child list uses the same hierarchical renderer. The old
 # single-child renderer silently discarded siblings or created duplicate
 # parent objects when child values changed. Do not use it for YANG lists.
 child_paths=metadata.get('child_paths',{})
 root_keys=[k for k in metadata.get('key_leaves',[]) if k not in child_paths]
 legacy_independent_rows=False
 if child_paths and root_keys and all(all(k in row for k in root_keys) for row in records):
  # Some older CIQs omit a parent key marker (e.g. a TAC) even though
  # it distinguishes separate parent objects. Keep these records separate
  # rather than coalescing incompatible parent scalar values.
  scalar_by_parent={}
  conflicting_parent=False
  for row in records:
   identity=tuple((k,row[k]) for k in root_keys)
   scalars={k:v for k,v in row.items() if k not in child_paths}
   old=scalar_by_parent.setdefault(identity,{})
   if any(k in old and old[k]!=v for k,v in scalars.items()):
    conflicting_parent=True
    break
   old.update(scalars)
  if not conflicting_parent:
   return render_multilevel_group(variable,records,source,metadata)
  if any(len(parts)>2 for parts in child_paths.values()):
   raise ValueError(f'{variable}: inconsistent parent values with multilevel child lists; verify declared YANG keys')
  # These are distinct parent records with an incomplete key declaration.
  legacy_independent_rows=True

 # Without a complete parent key, group identity cannot be proven. Preserve
 # legacy rendering for isolated records, but refuse ambiguous multiple-row
 # child data rather than silently dropping or combining instances.
 if child_paths and len(records)>1 and not legacy_independent_rows:
  raise ValueError(f'{variable}: child-list rows lack a complete declared parent key; cannot safely combine them')
 lines=['####################']
 path=metadata.get('yang_path') or (source.split(' | ')[1] if ' | ' in source else '')
 if path:lines.extend(_comment_lines('Xpath: '+path))
 cli=metadata.get('cli_command','')
 if cli:lines.extend(_comment_lines('CLI Commands: '+cli))
 desc=metadata.get('description','')
 if desc:
  lines.append('### Description:');lines.extend(_comment_lines(desc))
 deps=metadata.get('dependencies','')
 if deps:
  dep_lines=str(deps).splitlines()
  lines.extend(_comment_lines('Dependencies: '+dep_lines[0]))
  for extra in dep_lines[1:]:lines.extend(_comment_lines(extra))
 lines.extend(['####################','',f'{variable} = ['])
 string_leaves=set(metadata.get('string_leaves',[]))
 def typed_hcl(key,value,container=None):
  # Declared string types are indexed by complete YANG child path. A nested
  # leaf is rendered with its short name, but must retain its original type.
  qualified=f'{container}_{key}' if container else key
  return json.dumps(str(value),ensure_ascii=False) if qualified in string_leaves or key in string_leaves else hcl(value)
 child_paths=metadata.get('child_paths',{})
 for rec in records:
  lines.append('  {')
  scalar=[(k,v) for k,v in rec.items() if k not in child_paths]
  groups={}
  for k,v in rec.items():
   if k in child_paths:
    container,leaf=child_paths[k]
    groups.setdefault(container,[]).append((leaf,v))
  active=[(k,v) for k,v in scalar if not isinstance(v,DesignDefault)]
  defaults=[(k,v) for k,v in scalar if isinstance(v,DesignDefault)]
  width=max((len(k) for k,v in scalar),default=0)
  for k,v in active:lines.append(f'    {k.ljust(width)} = {typed_hcl(k,v)}')
  if defaults:
   lines.append('    ### Defaults used by design')
   for k,v in defaults:lines.append(f'    # {k.ljust(width)} = {typed_hcl(k,v)}')
  for container,leaves in groups.items():
   lines.append(f'    {container} = [')
   lines.append('      {')
   w=max(len(k) for k,v in leaves)
   for k,v in leaves:
    if not isinstance(v,DesignDefault):lines.append(f'        {k.ljust(w)} = {typed_hcl(k,v,container)}')
   defaults_child=[(k,v) for k,v in leaves if isinstance(v,DesignDefault)]
   if defaults_child:
    lines.append('        ### Defaults used by design')
    for k,v in defaults_child:lines.append(f'        # {k.ljust(w)} = {typed_hcl(k,v,container)}')
   lines.extend(['      },','    ]'])
  lines.extend(['    terrayang = {','      # with_defaults = "report-all"','    }','  },'])
 lines.extend([']',''])
 return '\n'.join(lines)

def write_group(output,variable,records,source,metadata=None):
 output.write_text(render_group(variable,records,source,metadata),encoding='utf-8')

def launch_gui(initial_workbook=None, initial_output=None):
 """Desktop launcher; importing Tk is deferred so headless CLI remains available."""
 try:
  import tkinter as tk
  from tkinter import ttk, filedialog, messagebox
 except ImportError as exc:
  print('Tkinter is unavailable. Install Python with Tk support or use CLI mode.', file=sys.stderr)
  return 3
 try:
  root=tk.Tk()
 except Exception as exc:
  print(f'Cannot open a desktop GUI ({exc}). Use CLI with a workbook path.',file=sys.stderr)
  return 3
 root.title('Terraform Script Generator')
 root.geometry('940x700');root.minsize(760,590)
 wb_var=tk.StringVar(value=str(initial_workbook or ''))
 out_var=tk.StringVar(value=str(initial_output or ''))
 mode=tk.StringVar(value='whole')
 combine_mode=tk.StringVar(value='separate')
 status=tk.StringVar(value='Select an Excel workbook, then click Scan Workbook.')
 source_cache={'sheets':[],'catalog':[],'warnings':[],'path':None}
 frame=ttk.Frame(root,padding=16);frame.pack(fill='both',expand=True)
 frame.columnconfigure(1,weight=1)
 ttk.Label(frame,text='Terraform Script Generator',font=('TkDefaultFont',16,'bold')).grid(row=0,column=0,columnspan=3,sticky='w',pady=(0,14))
 ttk.Label(frame,text='Input Excel workbook (.xlsx)').grid(row=1,column=0,sticky='w')
 ttk.Entry(frame,textvariable=wb_var).grid(row=1,column=1,sticky='ew',padx=8)
 def browse_wb():
  path=filedialog.askopenfilename(title='Select input workbook',filetypes=[('Excel workbook','*.xlsx'),('All files','*.*')])
  if path:wb_var.set(path);scan()
 ttk.Button(frame,text='Browse...',command=browse_wb).grid(row=1,column=2,sticky='ew')
 ttk.Label(frame,text='Output folder').grid(row=2,column=0,sticky='w',pady=8)
 ttk.Entry(frame,textvariable=out_var).grid(row=2,column=1,sticky='ew',padx=8)
 def browse_out():
  folder=filedialog.askdirectory(title='Select output folder')
  if folder:out_var.set(folder)
 ttk.Button(frame,text='Browse...',command=browse_out).grid(row=2,column=2,sticky='ew')
 ttk.Button(frame,text='Scan Workbook',command=lambda:scan()).grid(row=3,column=0,sticky='w',pady=8)
 group=ttk.LabelFrame(frame,text='Processing mode',padding=8);group.grid(row=4,column=0,columnspan=3,sticky='ew',pady=8)
 for i,(val,label) in enumerate([('whole','Whole workbook'),('sheets','Selected sheets'),('configs','Selected configurations'),('both','Selected sheets + configurations')]):
  ttk.Radiobutton(group,text=label,variable=mode,value=val).grid(row=i//2,column=i%2,sticky='w',padx=8,pady=3)
 ttk.Label(group,text='Output grouping:').grid(row=2,column=0,sticky='w',padx=8,pady=(9,2))
 ttk.Combobox(group,textvariable=combine_mode,state='readonly',values=('separate','sheet','all'),width=22).grid(row=2,column=1,sticky='w',padx=8,pady=(9,2))
 selectors=ttk.Frame(frame);selectors.grid(row=5,column=0,columnspan=3,sticky='nsew',pady=(6,8))
 selectors.columnconfigure(0,weight=1);selectors.columnconfigure(1,weight=1);selectors.rowconfigure(1,weight=1)
 frame.rowconfigure(5,weight=1)
 ttk.Label(selectors,text='Worksheets (Ctrl/Shift for multiple)').grid(row=0,column=0,sticky='w')
 ttk.Label(selectors,text='Configurations (Ctrl/Shift for multiple)').grid(row=0,column=1,sticky='w')
 sheet_list=tk.Listbox(selectors,selectmode='extended',exportselection=False,height=12)
 sheet_list.grid(row=1,column=0,sticky='nsew',padx=(0,8))
 config_list=tk.Listbox(selectors,selectmode='extended',exportselection=False,height=12)
 config_list.grid(row=1,column=1,sticky='nsew')
 names=[]
 def refresh_configs(event=None):
  previous={config_list.get(i) for i in config_list.curselection()}
  config_list.delete(0,'end');names.clear()
  selected={sheet_list.get(i) for i in sheet_list.curselection()}
  for c in source_cache['catalog']:
   if mode.get()=='both' and selected and c['sheet'] not in selected:continue
   label=f"{c['sheet']}  /  {c['config']}  ({len(c['entries'])} records)"
   names.append(c);config_list.insert('end',label)
   if label in previous:config_list.selection_set('end')
 sheet_list.bind('<<ListboxSelect>>',refresh_configs)
 def scan():
  path=Path(wb_var.get().strip().strip('"'))
  if not path.is_file():messagebox.showerror('Missing workbook','Choose an existing Excel workbook.');return
  try:sheets,catalog,warnings=discover(path)
  except Exception as e:messagebox.showerror('Workbook error',str(e));return
  source_cache.update(sheets=sheets,catalog=catalog,warnings=warnings,path=path.resolve())
  sheet_list.delete(0,'end')
  for x in sheets:sheet_list.insert('end',x)
  refresh_configs()
  if not out_var.get().strip():out_var.set(str(path.parent))
  status.set(f'Scanned {len(sheets)} worksheets / {len(catalog)} configurations / {len(warnings)} warnings.')
 def generate():
  path=Path(wb_var.get().strip().strip('"'))
  if source_cache['path'] is None or path.resolve()!=source_cache['path']:
   messagebox.showwarning('Scan required','Click Scan Workbook after selecting the input file.');return
  selected_sheets={sheet_list.get(i) for i in sheet_list.curselection()}
  selected_config_ids={id(names[i]) for i in config_list.curselection()}
  selected=[]
  for item in source_cache['catalog']:
   if mode.get() in ('sheets','both') and item['sheet'] not in selected_sheets:continue
   if mode.get() in ('configs','both') and id(item) not in selected_config_ids:continue
   selected.append(item)
  if not selected:
   messagebox.showwarning('Nothing selected','Select at least one matching worksheet/configuration, or choose Whole workbook.');return
  folder=out_var.get().strip().strip('"')
  if not folder:messagebox.showwarning('Output folder required','Choose a destination folder.');return
  try:
   dest=Path(folder).expanduser().resolve()
   count,records,dest=export_configs(path,dest,selected,source_cache['warnings'],[x for x in source_cache['sheets'] if x in selected_sheets] if mode.get() in ('sheets','both') else ['*'], [x['config'] for x in selected] if mode.get() in ('configs','both') else ['*'],combine=combine_mode.get())
  except Exception as e:messagebox.showerror('Generation error',str(e));return
  status.set(f'Generated {count} configurations ({records} records) into {dest}')
  messagebox.showinfo('Generation complete',f'Created {count} .tfvars files with {records} records.\n\nFolder: {dest}\nParser warnings: {len(source_cache["warnings"])} (see manifest.json).')
 ttk.Button(frame,text='Generate Terraform Files',command=generate).grid(row=6,column=0,columnspan=3,sticky='ew',pady=(8,8))
 ttk.Label(frame,textvariable=status,wraplength=850).grid(row=7,column=0,columnspan=3,sticky='w')
 ttk.Label(frame,text='Note: Outputs are generic HCL; verify variable schema and types against your Terraform modules.',wraplength=850).grid(row=8,column=0,columnspan=3,sticky='w',pady=(10,0))
 if wb_var.get() and Path(wb_var.get()).is_file():scan()
 root.mainloop()
 return 0

def timestamped_run_folder(base):
 """Create exactly one timestamped run folder under the selected destination.

 If an older version's default `Terraform_Output` directory is supplied,
 place the timestamped folder alongside it rather than inside it.
 """
 base=Path(base).expanduser().resolve()
 if base.name.casefold() == 'terraform_output':
  base=base.parent
 base.mkdir(parents=True,exist_ok=True)
 stamp=dt.datetime.now().strftime('%Y%m%d_%H%M%S')
 for suffix in range(1,10000):
  postfix='' if suffix == 1 else f'_{suffix}'
  folder=base/f'Terraform_Output_{stamp}{postfix}'
  try:
   folder.mkdir()
   return folder
  except FileExistsError:
   continue
 raise RuntimeError('Unable to create a unique timestamped output folder')

def combined_filename(workbook, selected, sheet_patterns=None, config_patterns=None):
 """Name a merged TFVARS file after what was selected; never use a generic name.

 A single configuration takes precedence. Explicit config choices use config names;
 explicit sheet choices use sheet names; full workbook uses the workbook name.
 Long or duplicate-name selections are shortened deterministically.
 """
 import hashlib
 sheets=list(dict.fromkeys(block['sheet'] for block in selected))
 configs=list(dict.fromkeys(block['config'] for block in selected))
 if len(selected)==1:
  names=[selected[0]['config']]
 elif config_patterns:
  names=configs
 elif sheet_patterns:
  names=sheets
 elif len(sheets)==1:
  names=sheets
 else:
  names=[workbook.stem]
 parts=[clean_name(x) for x in names if clean_name(x)]
 if not parts:parts=[clean_name(workbook.stem) or 'configuration']
 stem='_'.join(parts)
 if len(stem)>110:
  suffix=hashlib.sha256(stem.encode('utf-8')).hexdigest()[:10]
  stem=stem[:99].rstrip('_')+'_'+suffix
 return stem+'.tfvars'

def consolidate_same_yang_container(selected, warnings):
 """Combine CIQ sections defining the same canonical YANG list.

 Identity is (YANG path, Terraform variable), never the displayed worksheet
 title or an arbitrary numbered suffix. Duplicate parent/reference copies are
 ignored; distinct parent keys and repeated child rows are retained in order.
 """
 from collections import OrderedDict
 groups=OrderedDict()
 for block in selected:
  path=str(block.get('path','')).strip().rstrip('/')
  name=block['config']
  # A direct table has no unique schema; never infer a common container from it.
  key=(path,name) if path and path!='direct table' else ('direct',name,block['sheet'])
  if key not in groups:
   new=dict(block)
   new['entries']=[]
   new['metadata']=dict(block.get('metadata') or {})
   new['source_sheets']=[block['sheet']]
   new['_record_signatures']=set()
   groups[key]=new
  target=groups[key]
  if block['sheet'] not in target['source_sheets']:
   target['source_sheets'].append(block['sheet'])
  source_meta=block.get('metadata') or {}
  dest_meta=target['metadata']
  for field in ('parameter_order','key_leaves','string_leaves'):
   merged=list(dest_meta.get(field) or [])
   for leaf in source_meta.get(field) or []:
    if leaf not in merged:merged.append(leaf)
   if merged:dest_meta[field]=merged
  for leaf,parts in (source_meta.get('child_paths') or {}).items():
   dest_meta.setdefault('child_paths',{})
   if leaf in dest_meta['child_paths'] and dest_meta['child_paths'][leaf]!=parts:
    raise ValueError(f"{name}: inconsistent YANG child path for {leaf}")
   dest_meta['child_paths'][leaf]=parts
  for field in ('yang_path','cli_command','description'):
   existing=dest_meta.get(field)
   incoming=source_meta.get(field)
   if incoming and existing and str(incoming).strip()!=str(existing).strip():
    warnings.append(f"{name}: differing {field} across sections; retained first section's value")
   elif incoming and not existing:dest_meta[field]=incoming
  deps=list(dict.fromkeys(line.strip() for line in str(dest_meta.get('dependencies') or '').splitlines() if line.strip()))
  for line in str(source_meta.get('dependencies') or '').splitlines():
   line=line.strip()
   if line and line not in deps:deps.append(line)
  if deps:dest_meta['dependencies']='\n'.join(deps)
  for record in block['entries']:
   signature=json.dumps(record,sort_keys=True,default=str,ensure_ascii=False)
   if signature not in target['_record_signatures']:
    target['_record_signatures'].add(signature)
    target['entries'].append(record)
 for group in groups.values():
  group.pop('_record_signatures',None)
  group['entries']=reorder_entries(group['entries'],group['metadata'].get('parameter_order',[]))
 return list(groups.values())


def export_configs(workbook,out,selected,warnings,sheet_patterns=None,config_patterns=None,combine='separate'):
 """Export one canonical Terraform variable for each unique YANG container.

 Repeated sections are consolidated before output grouping. Where multiple
 sheets define the same canonical variable in sheet mode, the cross-sheet
 variable is emitted to its own named file to prevent duplicate variable
 assignments in multiple sheet output files.
 """
 import hashlib
 if combine not in ('separate','sheet','all'):
  raise ValueError(f'Unknown combine mode: {combine}')
 original_selected=selected
 selected=consolidate_same_yang_container(selected,warnings)
 # Preserve the relative order of every other configuration, but always emit
 # the consolidated AMF/MME network-element groups as the final block in any
 # output file that contains them. Stable sorting does not affect members.
 selected.sort(key=lambda block: block['config'] == 'amfmme_ne_grp_entry')
 # Plan and render in memory before creating files, so a conflicting YANG
 # child key cannot cause a partially written Terraform configuration.
 grouped={};manifest=[];names={}
 combined=combined_filename(workbook,original_selected,sheet_patterns,config_patterns) if combine=='all' else None
 for block in selected:
  stem=block['config']
  path=block.get('path','')
  identity=(path,stem)
  previous=names.get(stem)
  if previous is not None and previous!=identity:
   # Different schema under a coincidentally equal variable name: refuse
   # ambiguous Terraform assignments rather than adding arbitrary _2 suffixes.
   raise ValueError(f"Different YANG paths both require variable '{stem}': {previous[0]} and {path}. Resolve naming explicitly.")
  names[stem]=identity
  if combine=='all':file=combined
  elif combine=='sheet' and len(block.get('source_sheets',[]))==1:
   file=clean_name(block['sheet'])+'.tfvars'
  else:file=stem+'.tfvars'
  body=render_group(stem,block['entries'],f"{block['sheet']} | {path} | row {block['row']}",block.get('metadata'))
  for message in (block.get('metadata') or {}).get('validation_warnings',[]):
   if message not in warnings:warnings.append(message)
  grouped.setdefault(file,[]).append(body)
  manifest.append({'sheet':', '.join(block.get('source_sheets',[block['sheet']])),
                   'path':path,'variable':stem,'records':len(block['entries']),'file':file})
 out=timestamped_run_folder(out)
 for file,bodies in grouped.items():
  (out/file).write_text('\n'.join(bodies),encoding='utf-8')
 (out/'manifest.json').write_text(json.dumps({'input':workbook.name,'combine_mode':combine,'generated_files':list(grouped),'selected_sheets':sheet_patterns or ['*'],'selected_configurations':config_patterns or ['*'],'outputs':manifest,'warnings':warnings},indent=2),encoding='utf-8')
 with (out/'summary.csv').open('w',newline='',encoding='utf-8-sig') as f:
  w=csv.DictWriter(f,fieldnames=['sheet','path','variable','records','file']);w.writeheader();w.writerows(manifest)
 return len(grouped),sum(x['records'] for x in manifest),out

def main(argv=None):
 p=argparse.ArgumentParser(description='Terraform Script Generator — dynamic CIQ workbook to .tfvars converter')
 p.add_argument('workbook',type=Path,nargs='?',help='Source .xlsx file (optional; absent opens desktop app)')
 p.add_argument('--gui',action='store_true',help='Open desktop GUI (with optional prefilled workbook)')
 p.add_argument('--output',type=Path,default=None)
 p.add_argument('--sheets',help='Sheet names / wildcard patterns separated by commas')
 p.add_argument('--configs','--config',help='Configuration names / wildcard patterns separated by commas')
 p.add_argument('--combine',choices=('separate','sheet','all'),default='separate',help='Write separate configuration files, one file per sheet, or a named merged file based on the selected sheet/configuration/workbook')
 p.add_argument('--list',action='store_true',help='List detected sheets and configurations without writing files')
 p.add_argument('--all',action='store_true',help='Process all sheets/configurations without a menu')
 p.add_argument('--interactive',action='store_true',help='Force interactive selection menu')
 p.add_argument('--include-empty',action='store_true',help='Reserved for backwards compatibility')
 p.epilog='Terraform Script Generator | Select whole workbook, worksheets, or configurations.'
 args=p.parse_args(argv)
 if args.gui or (args.workbook is None and not (args.list or args.sheets or args.configs or args.all or args.interactive)):
  return launch_gui(args.workbook,args.output)
 print('Terraform Script Generator')
 if args.workbook is None:p.error('Workbook required in command-line mode. Run with no arguments to open the desktop app.')
 if not args.workbook.is_file():p.error('Input Excel workbook does not exist')
 sheets,catalog,warnings=discover(args.workbook)
 if not sheets:p.error('No worksheets readable in input workbook')
 if args.list:
  for sheet in sheets:
   print(f'\n[{sheet}]')
   entries=[v for v in catalog if v['sheet']==sheet]
   for item in entries:print(f"  {item['config']:<42} {len(item['entries'])} record(s) / row {item['row']}")
   if not entries:print('  No configuration detected')
  if warnings:print('\nParser warnings:\n'+ '\n'.join('  '+w for w in warnings))
  return 0
 sheet_patterns=csv_patterns(args.sheets)
 config_patterns=csv_patterns(args.configs)
 if args.interactive or (not args.all and not sheet_patterns and not config_patterns and sys.stdin.isatty()):
  picked=choose_indices('Select worksheets',sheets)
  if picked is not None:sheet_patterns=list(picked)
  options=list(dict.fromkeys(x['config'] for x in catalog if not sheet_patterns or matches(x['sheet'],sheet_patterns)))
  if options:
   picked=choose_indices('Select configurations (from chosen sheets)',options)
   if picked is not None:config_patterns=list(picked)
 selected=[c for c in catalog if (not sheet_patterns or matches(c['sheet'],sheet_patterns)) and (not config_patterns or matches(c['config'],config_patterns) or matches(c['path'],config_patterns))]
 if not selected:
  print('No matching configuration found. Run --list to see available names.',file=sys.stderr)
  return 2
 out=args.output or args.workbook.parent
 count,records,out=export_configs(args.workbook,out,selected,warnings,sheet_patterns,config_patterns,combine=args.combine)
 print(f'Created {count} output files containing {len(selected)} configurations / {records} records in {out}')
 print(f'Parser warnings: {len(warnings)} (see manifest.json)')
 return 0
if __name__=='__main__':sys.exit(main())

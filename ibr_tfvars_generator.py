#!/usr/bin/env python3
"""Terraform Script Generator — dynamic CIQ converter with verified horizontal and vertical target-block extraction. Python 3.10+, stdlib only.
Output is generic .auto.tfvars HCL maps/lists. Adapt root variable names to your Terraform module.
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
 v=str(v).strip().replace('/','_').replace('-','_').lower()
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
  if is_header(r):break
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
  status=' '.join(str(v) for j in range(max(1,pstart-3),pstart) for v in indexed.get(j,{}).values())
  if 'out of scope' in status.lower():continue
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
  status=' '.join(str(v) for j in range(max(1,pstart-3),pstart) for v in indexed.get(j,{}).values())
  if 'out of scope' in status.lower():continue
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

def find_vertical_target_values(section, prm):
 """Parse explicit two-column target blocks, never the parameter-description hints.

 The left-hand leaf column may include slash-separated child-list paths.
 Empty repeated target blocks (often the Current/alternate environment) are ignored.
 """
 tables=[]
 for pos,(rn,row) in enumerate(section):
  if rn<=prm or not any(re.search(r"\bTarget Values\s*\(",str(v),re.I) for v in row.values()):
   continue
  # A target heading normally occupies the column immediately above the key column.
  heading_cols=[c for c,v in row.items() if re.search(r"\bTarget Values\s*\(",str(v),re.I)]
  if not heading_cols:continue
  keycol=heading_cols[0]
  rec={}
  for rr,values in section[pos+1:]:
   if rr>rn+200:break
   if any(re.search(r'\b(Target Values\s*\(|PARAMATERS|PARAMETERS|YANG\s*Path\s*:)',str(v),re.I) for v in values.values()):break
   key=str(values.get(keycol,'')).strip()
   if not is_parameter_name(key):
    if rec and rr>rn+2:break
    continue
   # Target is the cell directly to the right of the named leaf. Other cells
   # represent review indicators or current values, not the requested target.
   raw=values.get(keycol+1)
   val=normalize(raw)
   if val is None:continue
   if isinstance(val,str) and re.match(r'^(refer to|see value below)',val,re.I):continue
   rec[clean_name(key)]=val
  if rec:tables.append((rn,rec))
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
  colset={c for c,_ in columns}
  for rr,values in section[header_at+1:]:
   if rr>section[header_at][0]+5000:break
   if any(re.search(r"\b(?:Target Values\s*\(|Current Values\s*\(|PARAMATERS|PARAMETERS|YANG\s*Path\s*:)",str(v),re.I) for v in values.values()):break
   if sum(clean_name(values.get(c,'')) in known for c in colset)>=2:break
   record={}
   for col,key in columns:
    val=normalize(values.get(col))
    if val is not None:record[key]=val
   # A valid record needs its first/key leaf; avoids standalone notes.
   if not record or columns[0][1] not in record:
    blank+=1
    if records and blank>=2:break
    continue
   records.append({key:record[key] for key in declared if key in record})
   blank=0
  if records:results.append((rn,records))
 return results

def convert_sheet(name,rows):
 """Prefer explicit vertical target-value blocks over description-matrix placeholders.

 Other established CIQ layouts continue using the legacy/sparse parsers.
 """
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
  metadata={'yang_path':path,'cli_command':'','description':'','dependencies':'','parameter_order':[],'child_paths':{}}
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
     if '/' in str(leaf):
      parts=[clean_name(part.strip()) for part in str(leaf).split('/') if part.strip()]
      if len(parts)==2:metadata['child_paths'][key]=parts
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

def discover(workbook):
 catalog=[]; errors=[]; sheets=[]
 for sheet,rows in read_workbook(workbook):
  sheets.append(sheet)
  blocks,warnings=convert_sheet(sheet,rows)
  if not blocks:
   blocks,extra=discover_key_value_table(sheet,rows)
   warnings=extra if blocks else warnings+extra
  metadata=collect_section_metadata(rows)
  for leaf,entries,path,rn in blocks:
   meta=metadata.get((path,rn),{})
   entries=reorder_entries(entries,meta.get('parameter_order',[]))
   catalog.append({'sheet':sheet,'config':clean_name(leaf),'entries':entries,'path':path,'row':rn,'metadata':meta})
  errors.extend(f'{sheet}: {w}' for w in warnings)
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

def render_group(variable,records,source,metadata=None):
 """Emit CIQ metadata, ordered scalar leaves and nested child-list entries."""
 metadata=metadata or {}
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
  for k,v in active:lines.append(f'    {k.ljust(width)} = {hcl(v)}')
  if defaults:
   lines.append('    ### Defaults used by design')
   for k,v in defaults:lines.append(f'    # {k.ljust(width)} = {hcl(v)}')
  for container,leaves in groups.items():
   lines.append(f'    {container} = [')
   lines.append('      {')
   w=max(len(k) for k,v in leaves)
   for k,v in leaves:
    if not isinstance(v,DesignDefault):lines.append(f'        {k.ljust(w)} = {hcl(v)}')
   defaults_child=[(k,v) for k,v in leaves if isinstance(v,DesignDefault)]
   if defaults_child:
    lines.append('        ### Defaults used by design')
    for k,v in defaults_child:lines.append(f'        # {k.ljust(w)} = {hcl(v)}')
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
  messagebox.showinfo('Generation complete',f'Created {count} .auto.tfvars files with {records} records.\n\nFolder: {dest}\nParser warnings: {len(source_cache["warnings"])} (see manifest.json).')
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

def export_configs(workbook,out,selected,warnings,sheet_patterns=None,config_patterns=None,combine='separate'):
 """Export separate files, one file per sheet, or one merged file for all selected sections.

 Preserve each configuration's own header and root variable; collisions are suffixed
 to keep the output valid when duplicate configuration names exist.
 """
 if combine not in ('separate','sheet','all'):
  raise ValueError(f'Unknown combine mode: {combine}')
 out=timestamped_run_folder(out)
 manifest=[];used=set();grouped={}
 for block in selected:
  stem=block['config'];n=2
  while stem in used:stem=f"{block['config']}_{n}";n+=1
  used.add(stem)
  if combine=='all':file='terraform_combined.auto.tfvars'
  elif combine=='sheet':file=clean_name(block['sheet'])+'.auto.tfvars'
  else:file=stem+'.auto.tfvars'
  body=render_group(stem,block['entries'],f"{block['sheet']} | {block['path']} | row {block['row']}",block.get('metadata'))
  grouped.setdefault(file,[]).append(body)
  manifest.append({'sheet':block['sheet'],'path':block['path'],'variable':stem,'records':len(block['entries']),'file':file})
 for file,bodies in grouped.items():
  (out/file).write_text('\n'.join(bodies),encoding='utf-8')
 (out/'manifest.json').write_text(json.dumps({'input':workbook.name,'combine_mode':combine,'generated_files':list(grouped),'selected_sheets':sheet_patterns or ['*'],'selected_configurations':config_patterns or ['*'],'outputs':manifest,'warnings':warnings},indent=2),encoding='utf-8')
 with (out/'summary.csv').open('w',newline='',encoding='utf-8-sig') as f:
  w=csv.DictWriter(f,fieldnames=['sheet','path','variable','records','file']);w.writeheader();w.writerows(manifest)
 return len(grouped),sum(x['records'] for x in manifest),out

def main(argv=None):
 p=argparse.ArgumentParser(description='Terraform Script Generator — dynamic CIQ workbook to .auto.tfvars converter')
 p.add_argument('workbook',type=Path,nargs='?',help='Source .xlsx file (optional; absent opens desktop app)')
 p.add_argument('--gui',action='store_true',help='Open desktop GUI (with optional prefilled workbook)')
 p.add_argument('--output',type=Path,default=None)
 p.add_argument('--sheets',help='Sheet names / wildcard patterns separated by commas')
 p.add_argument('--configs','--config',help='Configuration names / wildcard patterns separated by commas')
 p.add_argument('--combine',choices=('separate','sheet','all'),default='separate',help='Write separate configuration files, one combined file per sheet, or one combined file for the selected run')
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

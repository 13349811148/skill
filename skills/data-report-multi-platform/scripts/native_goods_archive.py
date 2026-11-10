"""Keep native pivots and user tabs; transplant only Artifact-authored cells/styles.

This packaging bridge computes no business values or formula caches. Existing
parts and cells stay native; all new/updated content comes from Artifact export.
"""
from copy import deepcopy
from pathlib import Path
import io,posixpath,re,zipfile,os
import xml.etree.ElementTree as ET

NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG='http://schemas.openxmlformats.org/package/2006/relationships'
Q=lambda n:f'{{{NS}}}{n}'


def sheet_paths(parts):
    rels={r.attrib['Id']:posixpath.normpath(posixpath.join('xl',r.attrib['Target'].lstrip('/'))) if not r.attrib['Target'].startswith('/') else r.attrib['Target'].lstrip('/')
          for r in ET.fromstring(parts['xl/_rels/workbook.xml.rels'])}
    return {s.attrib['name']:rels[s.attrib[f'{{{REL}}}id']] for s in ET.fromstring(parts['xl/workbook.xml']).find(Q('sheets'))}


class Styles:
    def __init__(self,original,authored):
        self.original=ET.fromstring(original);self.authored=ET.fromstring(authored);self.cache={}

    def table(self,root,name):
        result=root.find(Q(name))
        if result is None:
            result=ET.Element(Q(name),{'count':'0'})
            order=['numFmts','fonts','fills','borders','cellStyleXfs','cellXfs','cellStyles','dxfs','tableStyles','colors','extLst']
            position=next((i for i,c in enumerate(root) if c.tag.split('}')[-1] in order and order.index(c.tag.split('}')[-1])>order.index(name)),len(root))
            root.insert(position,result)
        return result

    def append(self,name,item):
        table=self.table(self.original,name);encoded=ET.tostring(item)
        for i,current in enumerate(table):
            if ET.tostring(current)==encoded:return i
        table.append(item);table.set('count',str(len(table)));return len(table)-1

    def numfmt(self,index):
        if index<164:return index
        source=next((n for n in self.table(self.authored,'numFmts') if int(n.attrib['numFmtId'])==index),None)
        if source is None:raise ValueError(f'Artifact自定义数字格式缺失: {index}')
        table=self.table(self.original,'numFmts')
        for n in table:
            if n.attrib['formatCode']==source.attrib['formatCode']:return int(n.attrib['numFmtId'])
        item=deepcopy(source);value=max([163,*[int(n.attrib['numFmtId']) for n in table]])+1
        item.set('numFmtId',str(value));table.append(item);table.set('count',str(len(table)));return value

    def map(self,name,index):
        key=name,index
        if key in self.cache:return self.cache[key]
        item=deepcopy(self.table(self.authored,name)[index])
        if name in ('cellXfs','cellStyleXfs'):
            for attr,table in [('fontId','fonts'),('fillId','fills'),('borderId','borders')]:
                if attr in item.attrib:item.set(attr,str(self.map(table,int(item.attrib[attr]))))
            if 'numFmtId' in item.attrib:item.set('numFmtId',str(self.numfmt(int(item.attrib['numFmtId']))))
            if name=='cellXfs' and 'xfId' in item.attrib:item.set('xfId',str(self.map('cellStyleXfs',int(item.attrib['xfId']))))
        result=self.append(name,item);self.cache[key]=result;return result


def column_number(value):
    n=0
    for c in value:n=n*26+ord(c)-64
    return n


def transplant(source,staged,payload):
    """Overlay mode appends rate cells; full mode replaces generated monthly tabs."""
    with zipfile.ZipFile(source) as z:
        original={i.filename:z.read(i.filename) for i in z.infolist() if not i.is_dir()}
    with zipfile.ZipFile(staged) as z:
        authored={i.filename:z.read(i.filename) for i in z.infolist() if not i.is_dir()}
    source_sheets=sheet_paths(original);new_sheets=sheet_paths(authored)
    styles=Styles(original['xl/styles.xml'],authored['xl/styles.xml'])
    shared=[]
    if 'xl/sharedStrings.xml' in authored:shared=list(ET.fromstring(authored['xl/sharedStrings.xml']))
    def convert(cell):
        cell=deepcopy(cell)
        if 's' in cell.attrib:cell.set('s',str(styles.map('cellXfs',int(cell.attrib['s']))))
        if cell.attrib.get('t')=='s':
            value=cell.find(Q('v'));si=shared[int(value.text)];cell.remove(value);cell.set('t','inlineStr')
            inline=ET.SubElement(cell,Q('is'))
            for child in si:inline.append(deepcopy(child))
        return cell
    if payload.get('marginOnly'):
        prefix=r'(?:[A-Za-z_][\w.-]*:)?'
        for edit in payload['rateEdits']:
            name=edit['name'];part=source_sheets[name];raw=original[part].decode('utf-8')
            tree=ET.fromstring(authored[new_sheets[name]])
            donors={c.attrib['r']:c for c in tree.findall(f'{Q("sheetData")}/{Q("row")}/{Q("c")}')}
            targets=set(edit['targets']);by_row={}
            for address in targets:
                if address not in donors:raise ValueError(f'Artifact净利率单元格缺失: {name}/{address}')
                number=int(re.search(r'\d+$',address)[0]);by_row.setdefault(number,[]).append((address,convert(donors[address])))
            for number,cells in sorted(by_row.items()):
                pattern=rf'<{prefix}row\b[^>]*\br="{number}"[^>]*(?:/>|>.*?</{prefix}row>)'
                found=re.search(pattern,raw,flags=re.S)
                if found:
                    row=found[0]
                    if row.endswith('/>'):
                        tag=re.match(r'<([^\s/>]+)',row)[1];row=row[:-2]+f'></{tag}>'
                    for address,cell in cells:
                        markup=ET.tostring(cell,encoding='unicode')
                        existing=rf'<{prefix}c\b[^>]*\br="{re.escape(address)}"[^>]*(?:/>|>.*?</{prefix}c>)'
                        if re.search(existing,row,flags=re.S):row=re.sub(existing,lambda m:markup,row,flags=re.S)
                        else:row=re.sub(rf'</{prefix}row>',lambda m:markup+m[0],row,count=1)
                    raw=raw[:found.start()]+row+raw[found.end():]
                else:
                    element=ET.Element(Q('row'),{'r':str(number)})
                    for _,cell in cells:element.append(cell)
                    markup=ET.tostring(element,encoding='unicode')
                    raw=re.sub(rf'</{prefix}sheetData>',lambda m:markup+m[0],raw,count=1)
            # The new rate is appended; existing source/pivot column positions stay fixed.
            col=column_number(edit['column'])
            old_cols=re.search(rf'<{prefix}cols\b[^>]*>.*?</{prefix}cols>',raw,flags=re.S)
            root_tag=re.search(rf'<{prefix}worksheet\b[^>]*>',raw)[0]
            context=' '.join(re.findall(r'xmlns(?::[\w.-]+)?="[^"]*"',root_tag))
            cols=ET.fromstring(f'<wrapper {context}>'+old_cols[0]+'</wrapper>')[0] if old_cols else ET.Element(Q('cols'))
            updated=[]
            for item in cols:
                start,end=int(item.attrib['min']),int(item.attrib['max'])
                if start<=col<=end:
                    if start<col:
                        left=deepcopy(item);left.set('max',str(col-1));updated.append(left)
                    if end>col:
                        right=deepcopy(item);right.set('min',str(col+1));updated.append(right)
                else:updated.append(item)
            donor_cols=tree.find(Q('cols'))
            donor=next((c for c in donor_cols or [] if int(c.attrib['min'])<=col<=int(c.attrib['max'])),None)
            item=deepcopy(donor) if donor is not None else ET.Element(Q('col'),{'width':'18','customWidth':'1'})
            item.set('min',str(col));item.set('max',str(col))
            if 'style' in item.attrib:item.set('style',str(styles.map('cellXfs',int(item.attrib['style']))))
            updated.append(item);cols[:]=sorted(updated,key=lambda c:int(c.attrib['min']))
            markup=ET.tostring(cols,encoding='unicode')
            raw=raw[:old_cols.start()]+markup+raw[old_cols.end():] if old_cols else re.sub(rf'<{prefix}sheetData\b[^>]*>',lambda m:markup+m[0],raw,count=1)
            dimension=re.search(rf'<{prefix}dimension\b[^>]*\bref="([^"]+)"[^>]*/>',raw)
            if dimension:
                bounds=dimension[1].split(':');start,last=bounds[0],bounds[-1]
                end_col=re.search(r'[A-Z]+',last)[0];end_row=int(re.search(r'\d+',last)[0])
                if column_number(end_col)<col:end_col=edit['column']
                end_row=max(end_row,*by_row)
                markup=ET.tostring(ET.Element(Q('dimension'),{'ref':f'{start}:{end_col}{end_row}'}),encoding='unicode')
                raw=raw[:dimension.start()]+markup+raw[dimension.end():]
            original[part]=raw.encode('utf-8')
    else:
        # Keep user-created sheets/pivots while updating generated month sheets.
        workbook=ET.fromstring(original['xl/workbook.xml']);sheets=workbook.find(Q('sheets'))
        rels=ET.fromstring(original['xl/_rels/workbook.xml.rels']);types=ET.fromstring(original['[Content_Types].xml'])
        added_month=False
        for part in payload['sheets']:
            name=part['name'];tree=ET.fromstring(authored[new_sheets[name]])
            for row in tree.findall(f'{Q("sheetData")}/{Q("row")}'):
                for cell in list(row):
                    if cell.tag==Q('c'):
                        i=list(row).index(cell);row.remove(cell);row.insert(i,convert(cell))
            for item in tree.findall(f'{Q("cols")}/{Q("col")}'):
                if 'style' in item.attrib:item.set('style',str(styles.map('cellXfs',int(item.attrib['style']))))
            target=source_sheets.get(name)
            if not target:
                added_month=True
                number=max([0,*[int(re.search(r'sheet(\d+)\.xml$',p)[1]) for p in source_sheets.values()]])+1
                target=f'xl/worksheets/sheet{number}.xml';rid='rIdGoods'+str(number)
                while any(r.attrib['Id']==rid for r in rels):rid+='x'
                sid=max([0,*[int(s.attrib['sheetId']) for s in sheets]])+1
                ET.SubElement(sheets,Q('sheet'),{'name':name,'sheetId':str(sid),f'{{{REL}}}id':rid})
                ET.SubElement(rels,f'{{{PKG}}}Relationship',{'Id':rid,'Type':REL+'/worksheet','Target':target.removeprefix('xl/')})
                ct=types.tag.split('}')[0]+'}Override'
                ET.SubElement(types,ct,{'PartName':'/'+target,'ContentType':'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'})
                source_sheets[name]=target
            original[target]=ET.tostring(tree,encoding='utf-8',xml_declaration=True)
        if added_month:
            original['xl/workbook.xml']=ET.tostring(workbook,encoding='utf-8',xml_declaration=True)
            original['xl/_rels/workbook.xml.rels']=ET.tostring(rels,encoding='utf-8',xml_declaration=True)
            # OPC package metadata requires an unprefixed root in the importer.
            def opc_xml(tree):
                raw=ET.tostring(tree,encoding='unicode')
                matched=re.match(r'<([^:> ]+):',raw)
                if not matched:return raw.encode('utf-8')
                prefix=matched[1]
                return raw.replace('xmlns:'+prefix+'=', 'xmlns=').replace('<'+prefix+':','<').replace('</'+prefix+':','</').encode('utf-8')
            original['xl/_rels/workbook.xml.rels']=opc_xml(rels)
            original['[Content_Types].xml']=opc_xml(types)
    original['xl/styles.xml']=ET.tostring(styles.original,encoding='utf-8',xml_declaration=True)
    temporary=Path(staged).with_suffix('.native.xlsx')
    with zipfile.ZipFile(temporary,'w',compression=zipfile.ZIP_DEFLATED) as z:
        for name,content in original.items():z.writestr(name,content)
    os.replace(temporary,staged)

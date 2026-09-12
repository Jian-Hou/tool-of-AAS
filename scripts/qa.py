"""Audit output against source values; never infer success from record counts alone."""
import argparse
import ast
import json
import math
import sys
import zipfile
from decimal import Decimal, InvalidOperation
from pathlib import Path

from conversion import (ValidationError, analyze, read_workbook, read_aasx, validate_model, atomic_json, safe_id)

PROJECT_DIR=Path(__file__).resolve().parent.parent

def audit(workbook, model, settings=None, bindings=None):
    report,context,settings,_=analyze(workbook,settings,bindings)
    problems=[]
    if report['errors']:problems.extend('Input error: '+i['message'] for i in report['errors'])
    try:validate_model(model)
    except ValidationError as exc:problems.append(str(exc))
    def problem(path,expected,actual):
        problems.append(f'{path}: expected {expected!r}, got {actual!r}')
    def eq(path,expected,actual,numeric=False):
        if expected in (None,'') and actual in (None,''):return
        if numeric:
            try:
                e,a=Decimal(str(expected)),Decimal(str(actual))
                if e.is_finite() and a.is_finite() and e==a:return
            except (InvalidOperation,ValueError):pass
        elif str(expected)==str(actual):return
        problem(path,expected,actual)
    def index(children,path):
        result={}
        for c in children:
            key=c.get('idShort')
            if key in result:problems.append(f'{path}: duplicate child {key}')
            result[key]=c
        return result
    def vals(node):return {x.get('idShort'):x.get('value') for x in node.get('value',[])}
    try:
        submodels=index(model.get('submodels',[]),'submodels')
        ad=index(submodels.get('AssemblyDefinition',{}).get('submodelElements',[]),'AssemblyDefinition')
        td=index(submodels.get('TechnicalData',{}).get('submodelElements',[]),'TechnicalData')
        components={}
        for comp in ad.get('Components',{}).get('value',[]):
            parts=index(comp.get('value',[]),'component')
            label=vals(parts.get('Identity',{})).get('Label')
            if label in components:problems.append(f'Duplicate component label: {label}')
            components[label]=parts
        source=context['sheets'];expected_labels={str(r['label']).strip() for r in source.get('components',[])}
        eq('Component set',sorted(expected_labels),sorted(str(k) for k in components))
        parameters=index(td.get('TechnicalProperties',{}).get('value',[]),'TechnicalProperties')
        eq('Parameter group set',sorted(safe_id(s) for s in expected_labels),sorted(str(k) for k in parameters))
        checked=0
        for row in source.get('components',[]):
            label=str(row['label']).strip();parts=components.get(label,{})
            identity=vals(parts.get('Identity',{}))
            for raw,target in [('assembly_id','ComponentId'),('label','Label'),('tag','AssetTag'),('type','ComponentType'),('shape_type','ShapeTypeCode')]:
                eq(label+'/Identity/'+target,row.get(raw),identity.get(target),raw=='assembly_id');checked+=1
            # Parse raw source tuples independently of the converter implementation.
            coords=list(ast.literal_eval(str(row['coord'])))
            placement=[]
            for part in str(row['placement']).split('|'):
                placement.extend(ast.literal_eval(part))
            q=placement[:4]
            if settings['coordinate_system']['QuaternionOrder']=='WXYZ':q=q[1:]+q[:1]
            for group,keys,expected in [('Position',('X','Y','Z'),coords),('Orientation',('Qx','Qy','Qz','Qw'),q),('BoundingBox',('XMin','YMin','ZMin','XMax','YMax','ZMax'),placement[4:])]:
                actual=vals(parts.get(group,{}))
                for key,value in zip(keys,expected):eq(label+'/'+group+'/'+key,value,actual.get(key),True);checked+=1
            for col,target in [('coord','SourceCoord'),('placement','SourcePlacement')]:eq(label+'/'+target,row[col],parts.get(target,{}).get('value'))
            family=str(row['type']).strip()
            param_row=next((p for p in source.get(family+'_instances',[]) if str(p['label']).strip()==label),{})
            eq(label+'/GeometryTypeCode',param_row.get('shape_type'),identity.get('GeometryTypeCode'))
            eq(label+'/SourceParameters',param_row.get('params'),parts.get('SourceParameters',{}).get('value'))
            expected_ref=[safe_id(param_row.get('shape_type',''))]
            actual_ref=parts.get('TypeDefinition',{}).get('value',{}).get('keys',[])
            if not actual_ref or actual_ref[-1].get('value')!=expected_ref[0]:problem(label+'/TypeDefinition',expected_ref,actual_ref)
            actual=vals(parameters.get(safe_id(label),{}));expected_keys=set()
            for pair in str(param_row.get('params','')).split(';'):
                if not pair.strip():continue
                key,raw=pair.split('=',1);key=__import__('re').sub('[^A-Za-z0-9_]','_',key.strip())
                values=raw.split(',')
                for i,value in enumerate(values):
                    dest=key if len(values)==1 else f'{key}_{i+1}';expected_keys.add(dest)
                    eq(label+'/Parameters/'+dest,value.strip(),actual.get(dest),True);checked+=1
            eq(label+'/Parameter keys',sorted(expected_keys),sorted(actual))
        joints={}
        for j in ad.get('Joints',{}).get('value',[]):
            v=vals(j);jid=str(v.get('JointId'))
            if jid in joints:problems.append('Duplicate JointId '+jid)
            joints[jid]=v
        eq('Joint ID set',sorted(str(int(Decimal(str(r['joint_id'])))) for r in source.get('joint_instances',[])),sorted(joints))
        for row in source.get('joint_instances',[]):
            jid=str(int(Decimal(str(row['joint_id']))));actual=joints.get(jid,{})
            for raw,target in [('joint_type','JointType'),('side1_id','Side1Component'),('side1_sub','Side1Feature'),('side2_id','Side2Component'),('side2_sub','Side2Feature')]:
                eq('Joint_'+jid+'/'+target,row.get(raw),actual.get(target));checked+=1
        catalog=index(ad.get('TypeCatalog',{}).get('value',[]),'TypeCatalog')
        for family,group in [('pipeline','PipelineTypes'),('elbow','ElbowTypes'),('blackbox','BlackboxTypes'),('tank','TankTypes')]:
            definitions=index(catalog.get(group,{}).get('value',[]),group)
            for row in source.get(family+'_types',[]):
                actual=vals(definitions.get(safe_id(row['shape_type']),{}))
                for col,target in [('shape_type','TypeCode'),('params_needed','RequiredParameters'),('geometric_description','GeometricDescription'),('count','SourceCount')]:
                    eq(group+'/'+str(row['shape_type'])+'/'+target,row.get(col),actual.get(target),col=='count');checked+=1
        eq('Source SHA256',workbook.source_sha256,vals(ad.get('SourceData',{})).get('SHA256'))
        eq('DataQuality.Status',report['status'],vals(ad.get('DataQuality',{})).get('Status'))
    except (KeyError,TypeError,ValueError,StopIteration,SyntaxError) as exc:
        problems.append('Unable to complete source data comparison: '+repr(exc));checked=0
    return {'status':'failed' if problems else report['status'],'data_match':not problems,'checked_fields':checked,
            'errors':problems,'source_report':report}

def main(argv=None):
    parser=argparse.ArgumentParser(description='Compare AASX fields with Excel and validate references and metamodel constraints.')
    parser.add_argument('aasx',nargs='?',type=Path,default=PROJECT_DIR/'output'/'Assembly001.aasx')
    parser.add_argument('--excel',type=Path,default=PROJECT_DIR/'source'/'final_result.xlsx')
    parser.add_argument('--config',type=Path)
    parser.add_argument('--strict',action='store_true')
    parser.add_argument('--report',type=Path)
    args=parser.parse_args(argv)
    try:
        workbook=read_workbook(args.excel)
        with zipfile.ZipFile(args.aasx) as z:
            config=json.loads(args.config.read_text(encoding='utf-8')) if args.config else json.loads(z.read('aasx/files/config.json')) if 'aasx/files/config.json' in z.namelist() else {}
            result=audit(workbook,read_aasx(args.aasx),config.get('settings'),config.get('bindings'))
            if 'aasx/files/source.xlsx' in z.namelist() and z.read('aasx/files/source.xlsx')!=workbook.source_bytes:
                result['errors'].append('The embedded source workbook differs from the Excel file being audited.');result['status']='failed';result['data_match']=False
        if args.strict and result['source_report']['warnings']:
            result['errors'].append('Strict mode does not allow incomplete information.');result['status']='failed'
    except (ValidationError,OSError,ValueError,KeyError,zipfile.BadZipFile) as exc:
        result={'status':'failed','data_match':False,'errors':[str(exc)]}
    if args.report:atomic_json(args.report,result)
    print(json.dumps(result,ensure_ascii=False,indent=2))
    return 2 if result['status']=='failed' else 0

if __name__=='__main__':raise SystemExit(main())

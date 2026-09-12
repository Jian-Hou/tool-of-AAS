"""Command line entry point for the shared, validated conversion pipeline."""
import argparse
import json
from pathlib import Path
from conversion import ValidationError, build_model, read_workbook, write_aasx, atomic_json, normalize_settings

PROJECT_DIR = Path(__file__).resolve().parent.parent

def main(argv=None):
    parser=argparse.ArgumentParser(description='Convert Excel assembly data to AASX. Incomplete data is explicitly marked as draft.')
    parser.add_argument('excel',nargs='?',type=Path,default=PROJECT_DIR/'source'/'final_result.xlsx')
    parser.add_argument('--output',type=Path)
    parser.add_argument('--config',type=Path,help='JSON configuration containing settings and bindings')
    parser.add_argument('--report',type=Path)
    parser.add_argument('--strict',action='store_true',help='Reject export if any information is incomplete')
    parser.add_argument('--check-only',action='store_true')
    parser.add_argument('--json',action='store_true')
    args=parser.parse_args(argv)
    try:
        config=json.loads(args.config.read_text(encoding='utf-8')) if args.config else {}
        if not isinstance(config,dict) or set(config)-{'settings','bindings'}:
            raise ValidationError('Configuration files support only settings and bindings.')
        settings=normalize_settings(config.get('settings'))
        workbook=read_workbook(args.excel)
        model,report=build_model(workbook,settings,config.get('bindings'),strict=args.strict)
        out=args.output or PROJECT_DIR/'output'/(settings['name']+'.aasx')
        if not args.check_only:write_aasx(model,out,workbook,report,settings,config.get('bindings'))
        result={'status':report['status'],'counts':report['counts'],'report':report}
        if not args.check_only:result['output']=str(out.resolve())
        if args.report:atomic_json(args.report,result)
        if args.json:print(json.dumps(result,ensure_ascii=False))
        else:
            print(('Validation completed' if args.check_only else 'Export completed') + (': draft; some information is incomplete.' if report['warnings'] else ': passed the implemented checks.'))
            print(f'Components: {report["counts"]["components"]}; joints: {report["counts"]["joints"]}; parameters: {report["counts"]["parameters"]}; issues to resolve: {len(report["warnings"])}.')
            if not args.check_only:print(out.resolve())
            for issue in report['warnings']:print(f'  [{issue["code"]}] {issue["sheet"]} {issue["row"] or ""} {issue["component"]}: {issue["message"]}')
        return 0
    except (ValidationError,OSError,json.JSONDecodeError) as exc:
        result={'status':'failed','error':str(exc),'report':getattr(exc,'report',None)}
        if args.report:atomic_json(args.report,result)
        print(json.dumps(result,ensure_ascii=False) if args.json else 'Conversion failed: '+str(exc))
        return 2

if __name__=='__main__':
    raise SystemExit(main())

import argparse, json, sys
from pathlib import Path
from .builder import build,update_charts
from .qa import render,audit,inspect

def main():
    ap=argparse.ArgumentParser(prog='academic_ppt',description='Editable, semantic-first scholarly PPTX authoring')
    sub=ap.add_subparsers(dest='cmd',required=True)
    b=sub.add_parser('build',help='Generate original slides, optionally insert into a base PPTX')
    b.add_argument('spec');b.add_argument('output');b.add_argument('--base');b.add_argument('--after',type=int)
    r=sub.add_parser('render');r.add_argument('pptx');r.add_argument('outdir');r.add_argument('--dpi',type=int,default=125)
    a=sub.add_parser('audit');a.add_argument('pptx');a.add_argument('--pdf');a.add_argument('--spec');a.add_argument('--json')
    p=sub.add_parser('inspect');p.add_argument('pptx');p.add_argument('--json')
    u=sub.add_parser('chart-update');u.add_argument('pptx');u.add_argument('edits_json');u.add_argument('output')
    args=ap.parse_args()
    try:
        if args.cmd=='build':v=build(args.spec,args.output,base=args.base,after=args.after)
        elif args.cmd=='render':v=render(args.pptx,args.outdir,args.dpi)
        elif args.cmd=='audit':
            spec=json.loads(Path(args.spec).read_text(encoding='utf8')) if args.spec else None
            v=audit(args.pptx,args.pdf,spec)
            if args.json:Path(args.json).write_text(json.dumps(v,ensure_ascii=False,indent=2),encoding='utf8')
        elif args.cmd=='inspect':
            v=inspect(args.pptx)
            if args.json:Path(args.json).write_text(json.dumps(v,ensure_ascii=False,indent=2),encoding='utf8')
        elif args.cmd=='chart-update':v={'updated_charts':update_charts(args.pptx,args.edits_json,args.output),'output':args.output}
        print(json.dumps(v,ensure_ascii=False,indent=2))
        if args.cmd=='audit' and v['error_count']:sys.exit(2)
    except Exception as exc:
        print(f'ERROR: {exc}',file=sys.stderr);sys.exit(1)
if __name__=='__main__':main()

"""Local web application: isolated projects, revision checks and safe uploads."""
import json
import os
import re
import secrets
import tempfile
import threading
import uuid
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file, session
from werkzeug.exceptions import HTTPException
from conversion import (ValidationError, read_workbook, analyze, build_model, write_aasx,
                        normalize_settings, normalize_bindings, atomic_json)
from schema import default_settings, default_bindings, mapping_contract

PROJECT_DIR=Path(__file__).resolve().parent.parent

def create_app(project_dir=None, testing=False):
    project_dir=Path(project_dir or PROJECT_DIR).resolve()
    state_dir=project_dir/'.state'
    state_dir.mkdir(parents=True,exist_ok=True)
    key_path=state_dir/'secret.key'
    if not key_path.exists():
        try:
            with key_path.open('x',encoding='ascii') as f:f.write(secrets.token_hex(32))
        except FileExistsError:pass
    application=Flask(__name__)
    application.config.update(SECRET_KEY=key_path.read_text(encoding='ascii'),TESTING=testing,
                              MAX_CONTENT_LENGTH=12*1024*1024,SESSION_COOKIE_HTTPONLY=True,
                              SESSION_COOKIE_SAMESITE='Strict',TRUSTED_HOSTS=['localhost','127.0.0.1','[::1]'])
    locks={};lock_guard=threading.Lock()

    def profile_id():
        pid=session.get('project_id')
        if not isinstance(pid,str) or not re.fullmatch('[0-9a-f]{32}',pid):
            pid=uuid.uuid4().hex;session['project_id']=pid;session.permanent=True
        return pid

    def folder():
        result=state_dir/profile_id();result.mkdir(exist_ok=True)
        return result

    def lock():
        with lock_guard:return locks.setdefault(profile_id(),threading.RLock())

    def load_state():
        path=folder()/'project.json'
        if path.exists():
            try:
                state=json.loads(path.read_text(encoding='utf-8'))
                state['settings']=normalize_settings(state.get('settings'))
                state['bindings']=normalize_bindings(state.get('bindings'))
                return state
            except (OSError,ValueError) as exc:
                raise ValidationError('Unable to read the saved project settings. Keep the configuration file and contact the maintainer; it will not be overwritten automatically.') from exc
        state={'revision':0,'csrf':secrets.token_urlsafe(32),'filename':None,'input':None,
               'source_sha256':None,'settings':default_settings(),'bindings':default_bindings()}
        atomic_json(path,state)
        return state

    def save_state(state):
        state['revision']+=1
        atomic_json(folder()/'project.json',state)

    def workbook_for(state):
        if not state['input']:raise ValidationError('Load an Excel file first.')
        name=state['input']
        if not isinstance(name,str) or not re.fullmatch(r'[0-9a-f]{32}\.xlsx',name):
            raise ValidationError('The saved input path is invalid.')
        workbook=read_workbook(folder()/name,source_name=state['filename'])
        if workbook.source_sha256!=state['source_sha256']:
            raise ValidationError('The saved input file changed outside this application. Load it again.')
        return workbook

    def state_response(state):
        result={k:state[k] for k in ('revision','csrf','filename','source_sha256','settings','bindings')}
        result['mapping_contract']=mapping_contract()
        result['sheets']={};result['report']=None
        if state['input']:
            workbook=workbook_for(state)
            for name,rows in workbook.sheets.items():
                result['sheets'][name]={'columns':workbook.headers[name],'row_count':len(rows),
                                       'preview':[{k:str(v) for k,v in row.items() if k!='_row'} for row in rows[:5]]}
            result['report']=analyze(workbook,state['settings'],state['bindings'])[0]
        return result

    def checked_revision(state, supplied):
        if type(supplied) is not int or supplied!=state['revision']:
            return jsonify(error='The project was changed in another tab. Refresh the project state before continuing.',code='stale_revision',revision=state['revision']),409
        return None

    def body():
        result=request.get_json(silent=True)
        if not isinstance(result,dict):raise ValidationError('The request body must be a JSON object.')
        return result

    @application.before_request
    def protect_writes():
        if request.method in ('POST','PUT','PATCH','DELETE'):
            origin=request.headers.get('Origin')
            if origin and origin.rstrip('/')!=request.host_url.rstrip('/'):
                return jsonify(error='Cross-origin requests are not allowed.'),403
            with lock():
                state=load_state()
                if not secrets.compare_digest(request.headers.get('X-CSRF-Token',''),state['csrf']):
                    return jsonify(error='Page verification has expired. Refresh the page and try again.'),403

    @application.after_request
    def response_headers(response):
        response.headers['Content-Security-Policy']="default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self'; object-src 'none'; base-uri 'none'; frame-ancestors 'none'"
        response.headers['X-Content-Type-Options']='nosniff'
        response.headers['Cache-Control']='no-store'
        return response

    @application.errorhandler(ValidationError)
    def validation_error(exc):
        return jsonify(error=str(exc),report=exc.report),422

    @application.errorhandler(HTTPException)
    def http_error(exc):
        message='The file exceeds the 12 MB limit.' if exc.code==413 else exc.description
        return jsonify(error=message),exc.code

    @application.errorhandler(Exception)
    def unexpected_error(exc):
        application.logger.exception('Application error')
        return jsonify(error='The operation failed. The previous project state has been preserved. Check the server log.'),500

    @application.get('/')
    def index():return render_template('index.html')

    @application.get('/api/state')
    def get_state():
        with lock():return jsonify(state_response(load_state()))

    @application.post('/api/excel/load')
    def load_excel():
        file=request.files.get('file')
        if not file or not file.filename:raise ValidationError('No Excel file was selected.')
        filename=file.filename
        if '/' in filename or '\\' in filename or ':' in filename or '\x00' in filename or filename in ('.','..'):
            raise ValidationError('The file name must not contain directory or path characters.')
        if Path(filename).suffix.lower()!='.xlsx':raise ValidationError('Only .xlsx is supported. Save .xls files as .xlsx first.')
        try:revision=int(request.form.get('revision',''))
        except ValueError:raise ValidationError('A valid page revision is required.')
        with lock():
            state=load_state()
            conflict=checked_revision(state,revision)
            if conflict:return conflict
            fd,tmp=tempfile.mkstemp(prefix='upload_',suffix='.xlsx',dir=folder());os.close(fd)
            try:
                file.save(tmp)
                workbook=read_workbook(tmp,source_name=filename)
                candidate=dict(state,input=uuid.uuid4().hex+'.xlsx',filename=filename,source_sha256=workbook.source_sha256)
                candidate['bindings']=default_bindings()
                analyze(workbook,candidate['settings'],candidate['bindings'])
                os.replace(tmp,folder()/candidate['input'])
                save_state(candidate)
                return jsonify(state_response(candidate))
            finally:Path(tmp).unlink(missing_ok=True)

    @application.post('/api/bindings/save')
    def save_bindings():
        data=body()
        with lock():
            state=load_state();conflict=checked_revision(state,data.get('revision'))
            if conflict:return conflict
            candidate=dict(state,bindings=normalize_bindings(data.get('bindings')))
            save_state(candidate)
            return jsonify(state_response(candidate))

    @application.post('/api/settings')
    def save_settings():
        data=body()
        with lock():
            state=load_state();conflict=checked_revision(state,data.get('revision'))
            if conflict:return conflict
            candidate=dict(state,settings=normalize_settings(data.get('settings')))
            save_state(candidate)
            return jsonify(state_response(candidate))

    @application.post('/api/project')
    def save_project():
        data=body()
        with lock():
            state=load_state();conflict=checked_revision(state,data.get('revision'))
            if conflict:return conflict
            candidate=dict(state,settings=normalize_settings(data.get('settings')),
                           bindings=normalize_bindings(data.get('bindings')))
            save_state(candidate)
            return jsonify(state_response(candidate))

    @application.post('/api/generate-aasx')
    def generate_aasx():
        data=body()
        if type(data.get('strict',False)) is not bool:raise ValidationError('strict must be a boolean.')
        with lock():
            state=load_state();conflict=checked_revision(state,data.get('revision'))
            if conflict:return conflict
            workbook=workbook_for(state)
            model,report=build_model(workbook,state['settings'],state['bindings'],strict=data.get('strict',False))
            token=uuid.uuid4().hex
            output=folder()/'outputs'/(token+'.aasx')
            write_aasx(model,output,workbook,report,state['settings'],state['bindings'])
            atomic_json(output.with_suffix('.json'),{'name':state['settings']['name'],'revision':state['revision']})
            return jsonify(status=report['status'],report=report,counts=report['counts'],download_url='/api/download/'+token,
                           filename=state['settings']['name']+'.aasx',revision=state['revision'])

    @application.get('/api/download/<token>')
    def download(token):
        if not re.fullmatch('[0-9a-f]{32}',token):return jsonify(error='File not found.'),404
        path=folder()/'outputs'/(token+'.aasx')
        if not path.exists() or not path.with_suffix('.json').exists():return jsonify(error='File not found.'),404
        meta=json.loads(path.with_suffix('.json').read_text(encoding='utf-8'))
        return send_file(path,as_attachment=True,download_name=meta['name']+'.aasx',mimetype='application/octet-stream')

    @application.get('/api/config')
    def get_config():
        with lock():
            state=load_state()
            response=jsonify(settings=state['settings'],bindings=state['bindings'])
            response.headers['Content-Disposition']='attachment; filename="aas_config.json"'
            return response

    @application.get('/api/preview')
    def preview():
        with lock():
            state=load_state()
            model,report=build_model(workbook_for(state),state['settings'],state['bindings'])
            return jsonify(model=model,report=report,revision=state['revision'])

    return application

def main():
    import argparse
    import webbrowser
    from werkzeug.serving import make_server
    parser=argparse.ArgumentParser(description='Local AAS conversion tool')
    parser.add_argument('--port',type=int,default=5000)
    parser.add_argument('--no-browser',action='store_true')
    parser.add_argument('--project-dir',type=Path,help='Local project directory (defaults to the AAS directory)')
    args=parser.parse_args()
    application=create_app(args.project_dir)
    try:
        server=make_server('127.0.0.1',args.port,application,threaded=True)
    except (OSError,SystemExit):
        print(f'Port {args.port} is unavailable. Selecting a free port; existing processes will remain running.')
        server=make_server('127.0.0.1',0,application,threaded=True)
    url=f'http://127.0.0.1:{server.server_port}'
    print('AAS tool is running at '+url,flush=True)
    if not args.no_browser:webbrowser.open(url)
    try:server.serve_forever()
    except KeyboardInterrupt:pass
    finally:server.server_close()

if __name__=='__main__':main()

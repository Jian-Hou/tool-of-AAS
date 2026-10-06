"""Web routes for the generic Excel-to-AAS mapper; state is kept per browser session next to the assembly project."""
import json
import os
import re
import tempfile
import uuid
from pathlib import Path

from flask import jsonify, render_template, request, send_file

import aas_mapper as m
from conversion import ValidationError, atomic_json, text

TEMPLATE_NAME = re.compile(r'[A-Za-z0-9_.\- ]{1,120}\.(aasx|json)', re.IGNORECASE)


def register(application, project_dir, folder, lock, body, csrf, prune):
    library = Path(project_dir) / 'aas_templates'
    tables_cache, package_cache = {}, {}

    def workdir():
        result = folder() / 'mapper'
        result.mkdir(exist_ok=True)
        return result

    def load():
        path = workdir() / 'state.json'
        if not path.exists():
            return {'revision': 0, 'excel': None, 'aas': None, 'rules': []}
        try:
            return json.loads(path.read_text(encoding='utf-8'))
        except (OSError, ValueError) as exc:
            raise ValidationError('Unable to read the saved mapping project.') from exc

    def save(state):
        state['revision'] += 1
        atomic_json(workdir() / 'state.json', state)

    def load_env(state):
        if not state['aas']:
            raise ValidationError('Open or create an AAS first.')
        return json.loads((workdir() / 'working.json').read_text(encoding='utf-8'))

    def shell_of(state):
        if not state['aas'] or not state['aas']['shell']:
            raise ValidationError('Select an asset administration shell first.')
        return state['aas']['shell']

    def tables(state):
        if not state['excel']:
            return None
        path = workdir() / state['excel']['file']
        key = (str(path), state['excel']['sha256'])
        if key not in tables_cache:
            result = m.read_tables(path, state['excel']['filename'])
            if result.source_sha256 != state['excel']['sha256']:
                raise ValidationError('The saved Excel file changed outside this application. Load it again.')
            if len(tables_cache) > 8:
                tables_cache.clear()
            tables_cache[key] = result
        return tables_cache[key]

    def response(state, extra=None):
        result = {'revision': state['revision'], 'csrf': csrf(), 'rules': state['rules'], 'excel': None, 'aas': None}
        workbook = tables(state)
        if workbook:
            # JSON objects are serialized with sorted keys, so the workbook's sheet order is sent separately.
            result['excel'] = {'filename': state['excel']['filename'], 'sheets': m.table_summary(workbook), 'order': list(workbook.sheets)}
        if state['aas']:
            env, info = load_env(state), state['aas']
            result['aas'] = {'filename': info['filename'], 'format': info['format'], 'opened': bool(info['source']),
                             'shells': m.shells(env), 'shell': info['shell'], 'tree': m.tree(env, info['shell']) if info['shell'] else [],
                             'existing_problems': info['baseline'][:20], 'existing_problem_count': len(info['baseline'])}
        result.update(extra or {})
        return jsonify(result)

    def stale(state, supplied):
        if type(supplied) is not int or supplied != state['revision']:
            return jsonify(error='The mapping project was changed in another tab. Refresh the page before continuing.',
                           code='stale_revision', revision=state['revision']), 409
        return None

    def uploaded(suffixes, needs_revision=True):
        file = request.files.get('file')
        if not file or not file.filename:
            raise ValidationError('No file was selected.')
        name = file.filename
        if any(ch in name for ch in '/\\:\x00') or name in ('.', '..'):
            raise ValidationError('The file name must not contain directory or path characters.')
        if Path(name).suffix.lower() not in suffixes:
            raise ValidationError('Unsupported file type. Expected: ' + ', '.join(suffixes))
        revision = None
        if needs_revision:
            try:
                revision = int(request.form.get('revision', ''))
            except ValueError:
                raise ValidationError('A valid page revision is required.') from None
        return file, name, revision

    def receive(file, suffix, directory):
        fd, tmp = tempfile.mkstemp(prefix='upload_', suffix=suffix, dir=directory)
        os.close(fd)
        file.save(tmp)
        return Path(tmp)

    def library_entries():
        library.mkdir(exist_ok=True)
        entries = []
        for path in sorted(library.iterdir(), key=lambda p: p.name.lower()):
            if not path.is_file() or not TEMPLATE_NAME.fullmatch(path.name):
                continue
            stat = path.stat()
            key = (path.name, stat.st_mtime_ns, stat.st_size)
            if key not in package_cache:
                try:
                    package_cache[key] = m.load_source(path)
                except ValidationError as exc:
                    package_cache[key] = exc
            package = package_cache[key]
            failed = isinstance(package, ValidationError)
            entries.append({'file': path.name, 'error': str(package) if failed else None,
                            'submodels': [] if failed else m.template_submodels(package.env)})
        return entries

    def library_package(name):
        name = text(name)
        if not TEMPLATE_NAME.fullmatch(name) or not (library / name).is_file():
            raise ValidationError('The template file was not found in the template library.')
        return m.load_source(library / name)

    def run(state):
        workbook = tables(state)
        if not workbook:
            raise ValidationError('Load an Excel file first.')
        if not state['rules']:
            raise ValidationError('Add at least one mapping first.')
        env, shell = load_env(state), shell_of(state)
        result, report = m.apply_rules(env, shell, workbook, state['rules'])
        if not report['errors']:
            known = set(state['aas']['baseline'])
            report['errors'] += [dict(rule=None, message='AAS validation: ' + failure, sheet='', row=None, column='')
                                 for failure in m.verify(result) if failure not in known]
        return env, shell, result, report

    @application.get('/mapper')
    def mapper_page():
        return render_template('mapper.html')

    @application.get('/api/mapper/state')
    def mapper_state():
        with lock():
            return response(load())

    @application.post('/api/mapper/excel')
    def mapper_excel():
        file, name, revision = uploaded(('.xlsx',))
        with lock():
            state = load()
            conflict = stale(state, revision)
            if conflict:
                return conflict
            tmp = receive(file, '.xlsx', workdir())
            try:
                workbook = m.read_tables(tmp, name)
                stored = uuid.uuid4().hex + '.xlsx'
                os.replace(tmp, workdir() / stored)
                previous = state['excel']
                state['excel'] = {'file': stored, 'filename': name, 'sha256': workbook.source_sha256}
                tables_cache[(str(workdir() / stored), workbook.source_sha256)] = workbook
                save(state)
                if previous:
                    (workdir() / previous['file']).unlink(missing_ok=True)
                return response(state)
            finally:
                tmp.unlink(missing_ok=True)

    def replace_aas(state, env, info):
        previous = state['aas']
        atomic_json(workdir() / 'working.json', env)
        state['aas'] = info
        save(state)
        if previous and previous['source'] and previous['source'] != info['source']:
            (workdir() / previous['source']).unlink(missing_ok=True)

    @application.post('/api/mapper/aas/open')
    def mapper_open():
        file, name, revision = uploaded(('.aasx', '.json'))
        with lock():
            state = load()
            conflict = stale(state, revision)
            if conflict:
                return conflict
            suffix = Path(name).suffix.lower()
            tmp = receive(file, suffix, workdir())
            try:
                package = m.load_source(tmp, name)
                found = m.shells(package.env)
                if not found:
                    raise ValidationError('This file contains no asset administration shell. Add it to the template library instead, or create a new AAS.')
                stored = 'source_' + uuid.uuid4().hex + suffix
                os.replace(tmp, workdir() / stored)
                replace_aas(state, package.env, {'filename': name, 'source': stored, 'spec_part': package.spec_part, 'format': package.spec_format,
                                                 'shell': found[0]['id'] if len(found) == 1 else None, 'baseline': m.verify(package.env)})
                return response(state)
            finally:
                tmp.unlink(missing_ok=True)

    @application.post('/api/mapper/aas/new')
    def mapper_new():
        data = body()
        with lock():
            state = load()
            conflict = stale(state, data.get('revision'))
            if conflict:
                return conflict
            env = m.new_environment(data.get('idShort'), data.get('globalAssetId', ''))
            shell = env['assetAdministrationShells'][0]
            replace_aas(state, env, {'filename': shell['idShort'] + '.aasx', 'source': None, 'spec_part': '', 'format': 'json',
                                     'shell': shell['id'], 'baseline': []})
            return response(state)

    @application.post('/api/mapper/aas/shell')
    def mapper_shell():
        data = body()
        with lock():
            state = load()
            conflict = stale(state, data.get('revision'))
            if conflict:
                return conflict
            m.get_shell(load_env(state), text(data.get('shell')))
            state['aas']['shell'] = text(data.get('shell'))
            save(state)
            return response(state)

    @application.post('/api/mapper/aas/edit')
    def mapper_edit():
        data = body()
        with lock():
            state = load()
            conflict = stale(state, data.get('revision'))
            if conflict:
                return conflict
            env, shell = load_env(state), shell_of(state)
            action = data.get('action')
            if action == 'submodel':
                m.add_submodel(env, shell, data.get('idShort'), data.get('semanticId', ''))
            elif action == 'template':
                m.add_template_submodel(env, shell, library_package(data.get('file')).env, text(data.get('submodel')), data.get('keepValues') is True)
            elif action == 'element':
                m.add_element(env, shell, m.parse_target(data.get('parent')), data.get('modelType'), data.get('idShort', ''),
                              data.get('valueType', ''), data.get('semanticId', ''))
            elif action == 'delete':
                m.delete_element(env, shell, m.parse_target(data.get('target')))
            else:
                raise ValidationError('Unknown structure change.')
            atomic_json(workdir() / 'working.json', m.canonical(env))
            save(state)
            return response(state)

    @application.get('/api/mapper/templates')
    def mapper_templates():
        return jsonify(templates=library_entries())

    @application.post('/api/mapper/templates')
    def mapper_add_template():
        file, name, _ = uploaded(('.aasx', '.json'), needs_revision=False)
        library.mkdir(exist_ok=True)
        tmp = receive(file, Path(name).suffix.lower(), library)
        try:
            package = m.load_source(tmp, name)
            if not package.env.get('submodels'):
                raise ValidationError('The template file contains no submodels.')
            stem = re.sub(r'[^A-Za-z0-9_.\-]', '_', Path(name).stem)[:100] or 'template'
            target, n = library / (stem + Path(name).suffix.lower()), 1
            while target.exists():
                n += 1
                target = library / f'{stem}_{n}{Path(name).suffix.lower()}'
            os.replace(tmp, target)
            return jsonify(templates=library_entries(), added=target.name)
        finally:
            tmp.unlink(missing_ok=True)

    @application.post('/api/mapper/rules')
    def mapper_rules():
        data = body()
        with lock():
            state = load()
            conflict = stale(state, data.get('revision'))
            if conflict:
                return conflict
            rules = data.get('rules')
            if isinstance(rules, dict):
                rules = rules.get('rules')
            state['rules'] = m.normalize_rules(rules)
            save(state)
            return response(state)

    @application.get('/api/mapper/rules.json')
    def mapper_rules_file():
        with lock():
            result = jsonify(format='excel-aas-mapping', version=1, rules=load()['rules'])
            result.headers['Content-Disposition'] = 'attachment; filename="aas_mapping.json"'
            return result

    @application.post('/api/mapper/check')
    def mapper_check():
        data = body()
        with lock():
            state = load()
            conflict = stale(state, data.get('revision'))
            if conflict:
                return conflict
            env, shell, result, report = run(state)
            return response(state, {'result': {'report': report, 'tree': m.compare_trees(m.tree(env, shell), m.tree(result, shell))}})

    @application.post('/api/mapper/import')
    def mapper_import():
        data = body()
        with lock():
            state = load()
            conflict = stale(state, data.get('revision'))
            if conflict:
                return conflict
            env, shell, result, report = run(state)
            if report['errors']:
                raise ValidationError('The import has errors; nothing was written.', report)
            info = state['aas']
            token = uuid.uuid4().hex
            output = workdir() / 'outputs' / (token + '.aasx')
            base = workdir() / info['source'] if info['source'] and info['spec_part'] else None
            m.write_package(output, result, base=base, package=m.Package(None, info['spec_part'], info['format'], info['filename']))
            filename = Path(info['filename']).stem + '.aasx'
            atomic_json(output.with_suffix('.json'), {'filename': filename})
            prune(output.parent, token)
            tree = m.compare_trees(m.tree(env, shell), m.tree(result, shell))
            # Later structure edits and imports continue from the imported state.
            atomic_json(workdir() / 'working.json', result)
            save(state)
            return response(state, {'result': {'report': report, 'tree': tree, 'download_url': '/api/mapper/download/' + token, 'filename': filename}})

    @application.get('/api/mapper/download/<token>')
    def mapper_download(token):
        if not re.fullmatch('[0-9a-f]{32}', token):
            return jsonify(error='File not found.'), 404
        path = workdir() / 'outputs' / (token + '.aasx')
        if not path.exists() or not path.with_suffix('.json').exists():
            return jsonify(error='File not found.'), 404
        meta = json.loads(path.with_suffix('.json').read_text(encoding='utf-8'))
        return send_file(path, as_attachment=True, download_name=meta['filename'], mimetype='application/octet-stream')

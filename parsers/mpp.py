from pathlib import Path
import pandas as pd
from .common import normalize

def _s(value): return None if value is None else str(value)
def _safe(task, method, default=None):
    try: return getattr(task, method)()
    except Exception: return default

def _owners(task):
    names=[]
    try:
        for assignment in task.getResourceAssignments():
            resource=assignment.getResource()
            if resource is not None and resource.getName() is not None: names.append(str(resource.getName()))
    except Exception: pass
    return ', '.join(dict.fromkeys(names)) or None

def _ancestors(task):
    result=[]; seen=set(); parent=_safe(task,'getParentTask')
    while parent is not None and id(parent) not in seen:
        seen.add(id(parent)); name=_safe(parent,'getName')
        if name: result.append(str(name))
        parent=_safe(parent,'getParentTask')
    return result

def _workstream(task):
    ancestors=_ancestors(task)
    # Highest meaningful ancestor below the project root is the delivery workstream.
    return ancestors[-2] if len(ancestors)>=2 else (ancestors[-1] if ancestors else None)

def _predecessors(task):
    values=[]
    try:
        for rel in task.getPredecessors():
            source=rel.getSourceTask()
            if source is not None: values.append(str(source.getID()))
    except Exception: pass
    return ', '.join(values) or None

def parse(path: Path):
    try:
        import jpype, mpxj
    except Exception as exc:
        raise RuntimeError('MPP support needs the Python package mpxj and Java 17+.') from exc
    try:
        if not jpype.isJVMStarted(): jpype.startJVM()
        from org.mpxj.reader import UniversalProjectReader
        project=UniversalProjectReader().read(str(path.resolve()))
    except Exception as exc:
        message=str(exc)
        if 'JVM' in message or 'java' in message.lower(): raise RuntimeError('Java could not be started. Install a 64-bit Java 17+ runtime/JDK.') from exc
        raise RuntimeError(f'Could not read this MPP file: {message}') from exc
    rows=[]
    for task in project.getTasks():
        name=_s(_safe(task,'getName'))
        if not name: continue
        try: progress=float(_safe(task,'getPercentageComplete',0) or 0)
        except Exception: progress=0
        rows.append({
            'ID':_safe(task,'getID'), 'Task':name, 'Start':_s(_safe(task,'getStart')), 'Finish':_s(_safe(task,'getFinish')),
            'Owner':_owners(task), '% Complete':progress, 'Milestone':bool(_safe(task,'getMilestone',False)),
            'Workstream':_workstream(task), 'Dependencies':_predecessors(task),
            'Outline Level':_safe(task,'getOutlineLevel'), 'Summary':bool(_safe(task,'getSummary',False))
        })
    if not rows: raise RuntimeError('No tasks were found in the MPP file.')
    return normalize(pd.DataFrame(rows))
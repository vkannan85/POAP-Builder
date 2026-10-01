from pathlib import Path
import pandas as pd
from .common import normalize

def _s(value):
    return None if value is None else str(value)

def _owners(task):
    names = []
    try:
        for assignment in task.getResourceAssignments():
            resource = assignment.getResource()
            if resource is not None and resource.getName() is not None:
                names.append(str(resource.getName()))
    except Exception:
        pass
    return ', '.join(dict.fromkeys(names)) or None

def _workstream(task):
    try:
        parent = task.getParentTask()
        if parent is not None and parent.getName() is not None:
            return str(parent.getName())
    except Exception:
        pass
    return None

def _predecessors(task):
    values = []
    try:
        for rel in task.getPredecessors():
            source = rel.getSourceTask()
            if source is not None:
                values.append(str(source.getID()))
    except Exception:
        pass
    return ', '.join(values) or None

def parse(path: Path):
    try:
        import jpype
        import mpxj
    except Exception as exc:
        raise RuntimeError('MPP support needs the Python package mpxj and Java 17+. Re-run setup, then install Java 17 or newer if this message remains.') from exc
    try:
        if not jpype.isJVMStarted():
            jpype.startJVM()
        from org.mpxj.reader import UniversalProjectReader
        project = UniversalProjectReader().read(str(path.resolve()))
    except Exception as exc:
        message = str(exc)
        if 'JVM' in message or 'java' in message.lower():
            raise RuntimeError('Java could not be started. Install a 64-bit Java 17+ runtime/JDK, close the portal, and start it again.') from exc
        raise RuntimeError(f'Could not read this MPP file: {message}') from exc
    rows = []
    for task in project.getTasks():
        name = _s(task.getName())
        if not name: continue
        try: progress = float(task.getPercentageComplete()) if task.getPercentageComplete() is not None else 0
        except Exception: progress = 0
        try: milestone = bool(task.getMilestone())
        except Exception: milestone = False
        rows.append({'Task':name,'Start':_s(task.getStart()),'Finish':_s(task.getFinish()),'Owner':_owners(task),'% Complete':progress,'Milestone':milestone,'Workstream':_workstream(task),'Dependencies':_predecessors(task)})
    if not rows: raise RuntimeError('No tasks were found in the MPP file.')
    return normalize(pd.DataFrame(rows))
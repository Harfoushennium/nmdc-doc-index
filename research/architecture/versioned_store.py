"""Research-only source-versioned SQLite store.
Immutable source versions; approval changes membership, never existing event payloads.
Not a production migration: flag, cache, custom-field, and configuration compatibility
requires separate verification before this store can become authoritative.
"""
from __future__ import annotations
import hashlib,json,sqlite3,uuid
from collections import defaultdict
from contextlib import contextmanager
from pathlib import Path
def canonical(value):
    return json.dumps(value,ensure_ascii=False,sort_keys=True,separators=(",",":"),default=str)
def checksum(value):
    return hashlib.sha256(canonical(value).encode("utf-8")).hexdigest()

SCHEMA="""
CREATE TABLE source_version(id TEXT PRIMARY KEY, source TEXT NOT NULL, metadata TEXT NOT NULL);
CREATE TABLE event(version TEXT REFERENCES source_version(id), event_key TEXT NOT NULL, doc TEXT NOT NULL, payload TEXT NOT NULL, PRIMARY KEY(version,event_key));
CREATE INDEX event_doc ON event(doc);
CREATE TABLE snapshot(id TEXT PRIMARY KEY, previous TEXT REFERENCES snapshot(id));
CREATE TABLE member(snapshot TEXT REFERENCES snapshot(id),source TEXT,version TEXT REFERENCES source_version(id),PRIMARY KEY(snapshot,source));
CREATE TABLE current(slot INTEGER PRIMARY KEY CHECK(slot=1),snapshot TEXT REFERENCES snapshot(id));
CREATE TABLE proposal(id TEXT PRIMARY KEY,base TEXT REFERENCES snapshot(id),state TEXT NOT NULL,revision INTEGER NOT NULL DEFAULT 0,flags TEXT NOT NULL);
CREATE TABLE change(proposal TEXT REFERENCES proposal(id),source TEXT,version TEXT REFERENCES source_version(id),accept INTEGER,note TEXT NOT NULL DEFAULT '',PRIMARY KEY(proposal,source));
"""
class Store:
    def __init__(self,path):
        self.db=sqlite3.connect(path,isolation_level=None)
        self.db.execute("PRAGMA foreign_keys=ON")
        self.db.execute("PRAGMA journal_mode=WAL")
        self.db.execute("PRAGMA synchronous=FULL")
        if not self.db.execute("SELECT 1 FROM sqlite_master WHERE name='snapshot'").fetchone():
            self.db.executescript(SCHEMA)
    @contextmanager
    def tx(self):
        self.db.execute("BEGIN IMMEDIATE")
        try:yield
        except BaseException:self.db.execute("ROLLBACK");raise
        else:self.db.execute("COMMIT")
    def current(self):
        r=self.db.execute("SELECT snapshot FROM current WHERE slot=1").fetchone()
        return r[0] if r else None
    def version(self,source,metadata,rows):
        rows=sorted(rows,key=lambda r:r["Event_Key"])
        if any(r.get("Source File")!=source or not r.get("Event_Key") for r in rows):raise ValueError("Invalid source/event identity")
        if len({r["Event_Key"] for r in rows})!=len(rows):raise ValueError("Duplicate within source")
        version=checksum([source,metadata,rows])
        if self.db.execute("SELECT 1 FROM source_version WHERE id=?",(version,)).fetchone():return version
        self.db.execute("INSERT INTO source_version VALUES(?,?,?)",(version,source,canonical(metadata)))
        self.db.executemany("INSERT INTO event VALUES(?,?,?,?)",((version,r["Event_Key"],str(r.get("Document No.","")),canonical(r)) for r in rows))
        return version
    def _validate(self,snapshot):
        duplicate=self.db.execute("""SELECT event_key FROM member m JOIN event e ON e.version=m.version
            WHERE snapshot=? GROUP BY event_key HAVING count(*)>1 LIMIT 1""",(snapshot,)).fetchone()
        if duplicate:raise ValueError("Cross-source duplicate identity")
    def import_legacy(self,manifest,rows):
        if self.current():raise ValueError("Import is into an empty store only")
        groups=defaultdict(list)
        for row in rows:groups[row["Source File"]].append(row)
        metadata={r["relative_path"]:r for r in manifest["files"]}
        if not set(groups)<=set(metadata):raise ValueError("Unmapped source")
        run=uuid.uuid4().hex
        with self.tx():
            self.db.execute("INSERT INTO snapshot VALUES(?,NULL)",(run,))
            for source,meta in metadata.items():
                v=self.version(source,meta,groups.get(source,[]))
                self.db.execute("INSERT INTO member VALUES(?,?,?)",(run,source,v))
            self._validate(run)
            self.db.execute("INSERT INTO current VALUES(1,?)",(run,))
        return run
    def stage(self,changes,flags=()):
        """changes map source to (metadata, records), or None for removal."""
        proposal=uuid.uuid4().hex
        with self.tx():
            base=self.current()
            if not base:raise ValueError("No approved baseline")
            self.db.execute("INSERT INTO proposal VALUES(?,?,'STAGED',0,?)",(proposal,base,canonical(list(flags))))
            for source,value in changes.items():
                v=self.version(source,*value) if value is not None else None
                self.db.execute("INSERT INTO change(proposal,source,version) VALUES(?,?,?)",(proposal,source,v))
        return proposal
    def save(self,proposal,revision,choices):
        with self.tx():
            r=self.db.execute("SELECT base,state,revision FROM proposal WHERE id=?",(proposal,)).fetchone()
            if not r or r!=(self.current(),"STAGED",revision):raise ValueError("Stale proposal or review")
            sources={r[0] for r in self.db.execute("SELECT source FROM change WHERE proposal=?",(proposal,))}
            if set(choices)!=sources or any(type(v[0]) is not bool or not isinstance(v[1],str) for v in choices.values()):raise ValueError("Incomplete or invalid choices")
            for source,(accept,note) in choices.items():self.db.execute("UPDATE change SET accept=?,note=? WHERE proposal=? AND source=?",(int(accept),note,proposal,source))
            self.db.execute("UPDATE proposal SET revision=revision+1 WHERE id=?",(proposal,))
        return revision+1
    def approve(self,proposal,revision,interrupt=False):
        run=uuid.uuid4().hex
        with self.tx():
            r=self.db.execute("SELECT base,state,revision,flags FROM proposal WHERE id=?",(proposal,)).fetchone()
            if not r or r[:3]!=(self.current(),"STAGED",revision) or revision==0:raise ValueError("Stale/unsaved approval")
            base=r[0]
            if json.loads(r[3]):raise ValueError("Flags require review")
            if self.db.execute("SELECT 1 FROM change WHERE proposal=? AND accept IS NULL",(proposal,)).fetchone():raise ValueError("Unsaved choice")
            self.db.execute("INSERT INTO snapshot VALUES(?,?)",(run,base))
            self.db.execute("INSERT INTO member SELECT ?,source,version FROM member WHERE snapshot=?",(run,base))
            for source,version in self.db.execute("SELECT source,version FROM change WHERE proposal=? AND accept=1",(proposal,)).fetchall():
                self.db.execute("DELETE FROM member WHERE snapshot=? AND source=?",(run,source))
                if version:self.db.execute("INSERT INTO member VALUES(?,?,?)",(run,source,version))
            self._validate(run)
            if interrupt:raise RuntimeError("Simulated interrupted transaction")
            self.db.execute("UPDATE current SET snapshot=? WHERE slot=1",(run,))
            self.db.execute("UPDATE proposal SET state='APPROVED' WHERE id=?",(proposal,))
        return run
    def rows(self,snapshot=None):
        snapshot=snapshot or self.current()
        rows=[json.loads(r[0]) for r in self.db.execute("SELECT e.payload FROM member m JOIN event e ON e.version=m.version WHERE m.snapshot=?",(snapshot,))]
        return sorted(rows,key=lambda r:tuple(str(r.get(k,"")) for k in ("Project No.","Document No.","Revision","Event_Key","Source File","Source Row")))
    def metadata(self,snapshot=None):
        return {r[0]:json.loads(r[1]) for r in self.db.execute("SELECT m.source,v.metadata FROM member m JOIN source_version v ON v.id=m.version WHERE m.snapshot=?",(snapshot or self.current(),))}
    def undo(self):
        with self.tx():
            r=self.db.execute("SELECT previous FROM snapshot WHERE id=?",(self.current(),)).fetchone()
            if not r or not r[0]:raise ValueError("No previous snapshot")
            self.db.execute("UPDATE current SET snapshot=? WHERE slot=1",(r[0],))
        return r[0]

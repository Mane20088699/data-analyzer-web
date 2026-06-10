#!/usr/bin/env python3
"""Data Analyzer — Web Edition  (Flask + Tailwind CSS, localhost)"""

# ── SSL bypass ────────────────────────────────────────────────────────────────
import ssl as _ssl
_ssl._create_default_https_context = _ssl._create_unverified_context

import httpx as _httpx
_orig_client = _httpx.Client.__init__
def _patched_client(self, *a, **kw): kw.setdefault("verify", False); _orig_client(self, *a, **kw)
_httpx.Client.__init__ = _patched_client
_orig_async = _httpx.AsyncClient.__init__
def _patched_async(self, *a, **kw): kw.setdefault("verify", False); _orig_async(self, *a, **kw)
_httpx.AsyncClient.__init__ = _patched_async

import requests as _req, urllib3 as _u3
_u3.disable_warnings()
_orig_req = _req.Session.request
def _patched_req(self, *a, **kw): kw.setdefault("verify", False); return _orig_req(self, *a, **kw)
_req.Session.request = _patched_req

# ── Standard imports ──────────────────────────────────────────────────────────
import os, re, sys, json, uuid, warnings, tempfile, zipfile, io
from pathlib import Path
from collections import Counter
from typing import List, Dict, Tuple, Optional
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
import networkx as nx
import nltk
import pdfplumber
import spacy
from docx import Document as DocxDocument
from keybert import KeyBERT
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from flask import Flask, request, jsonify, send_file, render_template, Response

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ── NLTK bootstrap ────────────────────────────────────────────────────────────
for _r in ["punkt", "punkt_tab", "stopwords"]:
    nltk.download(_r, quiet=True)
from nltk.corpus import stopwords as _sw
from nltk.tokenize import sent_tokenize
STOPWORDS = set(_sw.words("english"))

# ── Load local models ─────────────────────────────────────────────────────────
print("Loading NLP models…")
NLP: Optional[spacy.Language] = None
NLP_MODEL = "none"
for _m in ["en_core_web_trf", "en_core_web_lg", "en_core_web_md", "en_core_web_sm"]:
    try:
        NLP = spacy.load(_m); NLP_MODEL = _m
        print(f"  [OK] spaCy: {_m}"); break
    except OSError:
        pass
if NLP is None:
    os.system(f'"{sys.executable}" -m spacy download en_core_web_lg')
    NLP = spacy.load("en_core_web_lg"); NLP_MODEL = "en_core_web_lg"
NLP.max_length = 2_000_000

KW_MODEL = KeyBERT("all-MiniLM-L6-v2")
print("  [OK] KeyBERT: all-MiniLM-L6-v2\nReady.\n")

# ── Patterns ──────────────────────────────────────────────────────────────────
PROBLEM_WORDS = {
    "fail","fails","failed","failure","error","errors","bug","bugs",
    "defect","issue","issues","problem","problems","crash","crashes",
    "broken","corrupted","conflict","conflicts","incompatible","mismatch",
    "missing","undefined","exception","deadlock","leak","overflow",
    "risk","hazard","unsafe","vulnerable","vulnerability","violation",
    "breach","invalid","incorrect","wrong","timeout","blocked","frozen",
    "deprecated","obsolete","flaw","weakness",
}

CAUSAL_PATTERNS: List[Tuple[str, str]] = [
    (r"(\w[\w\s]{1,40}?)\s+causes?\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "causes"),
    (r"(\w[\w\s]{1,40}?)\s+leads?\s+to\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "leads to"),
    (r"(\w[\w\s]{1,40}?)\s+results?\s+in\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "results in"),
    (r"(\w[\w\s]{1,40}?)\s+triggers?\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "triggers"),
    (r"(\w[\w\s]{1,40}?)\s+depends?\s+on\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "depends on"),
    (r"(\w[\w\s]{1,40}?)\s+requires?\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "requires"),
    (r"(\w[\w\s]{1,40}?)\s+affects?\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "affects"),
    (r"(\w[\w\s]{1,40}?)\s+impacts?\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "impacts"),
    (r"(\w[\w\s]{1,40}?)\s+prevents?\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "prevents"),
    (r"(\w[\w\s]{1,40}?)\s+enables?\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "enables"),
    (r"(\w[\w\s]{1,40}?)\s+blocks?\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "blocks"),
    (r"if\s+(\w[\w\s]{1,40}?),?\s+then\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "if → then"),
    (r"(\w[\w\s]{1,40}?)\s+is\s+caused\s+by\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "caused by"),
    (r"(\w[\w\s]{1,40}?)\s+is\s+due\s+to\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "due to"),
    (r"(\w[\w\s]{1,40}?)\s+is\s+associated\s+with\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "associated with"),
    (r"(\w[\w\s]{1,40}?)\s+represents?\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "represents"),
    (r"(\w[\w\s]{1,40}?)\s+symbolizes?\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "symbolizes"),
    (r"(\w[\w\s]{1,40}?)\s+explores?\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "explores"),
    (r"(\w[\w\s]{1,40}?)\s+faces?\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "faces"),
    (r"(\w[\w\s]{1,40}?)\s+carries?\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "carries"),
    (r"(\w[\w\s]{1,40}?)\s+reflects?\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "reflects"),
    (r"(\w[\w\s]{1,40}?)\s+reveals?\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "reveals"),
    (r"(\w[\w\s]{1,40}?)\s+shows?\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "shows"),
    (r"(\w[\w\s]{1,40}?)\s+demands?\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "demands"),
    (r"(\w[\w\s]{1,40}?)\s+extracts?\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "extracts"),
    (r"(\w[\w\s]{1,40}?)\s+share[sd]?\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "shares"),
    (r"(\w[\w\s]{1,40}?)\s+enters?\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "enters"),
    (r"(\w[\w\s]{1,40}?)\s+reclaims?\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "reclaims"),
    (r"(\w[\w\s]{1,40}?)\s+recognizes?\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "recognizes"),
    (r"(\w[\w\s]{1,40}?)\s+escapes?\s+(\w[\w\s]{1,40}?)(?=[,.\n]|$)", "escapes"),
]

VARIABLE_PATTERNS: List[Tuple[str, str]] = [
    (r"\b([A-Z][a-zA-Z0-9_]{1,20})\s*=\s*([\d.]+\s*[a-zA-Z%°/]*)", "assignment"),
    (r"\b([a-zA-Z_]\w{1,20})\s+is\s+set\s+to\s+([\d.]+\s*[a-zA-Z%°/]*)", "configured"),
    (r"\b([a-zA-Z_]\w{1,20})\s*:\s*([\d.]+\s*[a-zA-Z%°/]*)\b", "defined"),
    (r"\b(threshold|limit|maximum|minimum|rate|value|count|size|capacity"
     r"|timeout|interval|budget|score|ratio|percentage|speed|temperature"
     r"|pressure|voltage|frequency|duration|weight|height|width|length)"
     r"\s+(?:of\s+[\w\s]{1,20}\s+)?(?:is\s+)?([\d.]+\s*(?:[a-zA-Z%°/]+)?)", "parameter"),
    (r"\b([\d.]+)\s*(percent|%|degrees?|ms|milliseconds?|seconds?|minutes?"
     r"|hours?|days?|meters?|km|kg|lb|MB|GB|TB|KB|RPM|fps|Hz|kHz|MHz|V|A|W)\b", "measurement"),
]

# ── Document readers ──────────────────────────────────────────────────────────
def read_pdf(path: str) -> str:
    parts = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t: parts.append(t)
    raw = "\n\n".join(parts)
    raw = re.sub(r'(?<!\n)\n(?!\n)', ' ', raw)
    raw = re.sub(r' {2,}', ' ', raw)
    return raw.strip()

def read_docx(path: str) -> str:
    doc = DocxDocument(path)
    parts = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            r = " | ".join(c.text.strip() for c in row.cells if c.text.strip())
            if r: parts.append(r)
    return "\n".join(parts)

def read_text(path: str) -> str:
    for enc in ["utf-8", "utf-8-sig", "latin-1", "cp1252"]:
        try: return Path(path).read_text(encoding=enc)
        except (UnicodeDecodeError, LookupError): pass
    return Path(path).read_bytes().decode("ascii", errors="replace")

def read_document(path: str) -> Tuple[str, str]:
    ext = Path(path).suffix.lower()
    if ext == ".pdf": return read_pdf(path), "PDF"
    if ext in (".docx", ".doc"): return read_docx(path), "Word Document"
    return read_text(path), "Text / ASCII"

# ── NLP helpers ───────────────────────────────────────────────────────────────
def chunk_text(text: str, max_chunk: int = 80_000) -> List[str]:
    if len(text) <= max_chunk: return [text]
    sents = sent_tokenize(text)
    chunks, cur = [], ""
    for s in sents:
        if len(cur) + len(s) > max_chunk and cur:
            chunks.append(cur.strip()); cur = s
        else:
            cur += " " + s
    if cur.strip(): chunks.append(cur.strip())
    return chunks or [text[:max_chunk]]

def _clean(text: str, n: int = 200) -> str:
    return re.sub(r"\s+", " ", text).strip()[:n]

# ── Extractors ────────────────────────────────────────────────────────────────
_PERSON_PREFIXES = re.compile(
    r"^(O'|O’|Mc|Mac|De|Di|Van|Von|Le|La|St\.?\s)\w", re.IGNORECASE)
_TRIVIAL_CONCEPTS = {
    "one","two","three","four","five","six","seven","eight","nine","ten",
    "first","second","third","fourth","fifth","last","next","new","old",
    "good","great","well","also","even","just","like","much","many","every",
    "said","say","told","tell","got","get","know","think","want","need",
    "go","come","see","look","take","give","make","use","find","turn",
    "day","time","man","woman","people","way","thing","place","part","year",
    "yes","no","not","yet","still","back","away","down","up","out","over",
    "never","always","already","once","again","around","together",
}

def extract_entities(text: str) -> List[Dict]:
    counter: Counter = Counter()
    for chunk in chunk_text(text):
        doc = NLP(chunk)
        for ent in doc.ents:
            t = ent.text.strip()
            label = ent.label_
            if len(t) <= 1: continue
            if label == "GPE" and (_PERSON_PREFIXES.match(t) or "'" in t):
                label = "PERSON"
            counter[(t, label)] += 1
    return [{"Entity": t, "Type": l, "Description": spacy.explain(l) or l, "Occurrences": c}
            for (t, l), c in sorted(counter.items(), key=lambda x: -x[1])]

def extract_concepts(text: str, top_n: int = 30) -> List[Dict]:
    try:
        kws = KW_MODEL.extract_keywords(text[:60_000], keyphrase_ngram_range=(1, 3),
                                         stop_words="english", top_n=top_n,
                                         use_mmr=True, diversity=0.5)
        results = []
        for kw, sc in kws:
            words = kw.lower().split()
            if len(words) == 1 and words[0] in _TRIVIAL_CONCEPTS: continue
            if re.fullmatch(r'[\d\s,.\-]+', kw): continue
            if len(words) == 1 and len(kw) <= 2: continue
            results.append({"Concept": kw, "Relevance Score": round(sc, 4)})
            if len(results) >= 25: break
        return results
    except Exception as exc:
        return [{"Concept": f"Error: {exc}", "Relevance Score": 0.0}]

def extract_variables(text: str) -> List[Dict]:
    results, seen = [], set()
    for sent in sent_tokenize(text):
        for pat, vtype in VARIABLE_PATTERNS:
            for m in re.finditer(pat, sent, re.IGNORECASE):
                try:
                    g = m.groups()
                    name, val = (g[0].strip(), g[1].strip()) if len(g) >= 2 else (None, None)
                    if not name or not val or len(name) < 2: continue
                    key = (name.lower()[:30], val[:20])
                    if key not in seen:
                        seen.add(key)
                        results.append({"Variable / Parameter": name, "Value": val,
                                        "Type": vtype, "Context": _clean(sent)})
                except Exception: continue
    return results

_SVO_DEPS = {"ROOT", "relcl", "xcomp", "advcl", "conj", "acl", "ccomp", "pcomp"}

def extract_svo(text: str) -> List[Dict]:
    results, seen = [], set()
    for chunk in chunk_text(text, 40_000):
        doc = NLP(chunk)
        for sent in doc.sents:
            for tok in sent:
                if tok.pos_ != "VERB" or tok.dep_ not in _SVO_DEPS: continue
                verb = tok.lemma_.lower()
                subjs = [t for t in tok.lefts  if t.dep_ in ("nsubj","nsubjpass","csubj")]
                objs  = [t for t in tok.rights if t.dep_ in ("dobj","attr","acomp")]
                for prep in tok.rights:
                    if prep.dep_ == "prep":
                        objs += [c for c in prep.children if c.dep_ == "pobj"]
                for s in subjs:
                    st = _clean(" ".join(t.text for t in s.subtree
                                        if t.dep_ not in ("det","punct") and len(t.text) > 1))
                    for o in objs:
                        ot = _clean(" ".join(t.text for t in o.subtree
                                             if t.dep_ not in ("det","punct") and len(t.text) > 1))
                        key = (st.lower()[:40], verb[:20], ot.lower()[:40])
                        if key not in seen and st and ot and len(st) > 2:
                            seen.add(key)
                            results.append({"Subject": st, "Relationship": verb, "Object": ot,
                                            "Rel. Type": "syntactic (SVO)",
                                            "Source Sentence": _clean(sent.text, 250)})
    return results[:200]

def extract_causal(text: str) -> List[Dict]:
    results, seen = [], set()
    for sent in sent_tokenize(text):
        for pat, rel in CAUSAL_PATTERNS:
            for m in re.finditer(pat, sent, re.IGNORECASE):
                try:
                    s, o = m.group(1).strip(), m.group(2).strip()
                    key = (s.lower()[:40], rel, o.lower()[:40])
                    if key not in seen and len(s) > 2 and len(o) > 2:
                        seen.add(key)
                        results.append({"Subject": s, "Relationship": rel, "Object": o,
                                        "Rel. Type": "causal", "Source Sentence": _clean(sent, 250)})
                except Exception: continue
    return results[:200]

def identify_problems(rels: List[Dict], entities: List[Dict], text: str) -> List[Dict]:
    problems = []
    ent_names = {e["Entity"].lower() for e in entities}

    G = nx.DiGraph()
    for r in rels:
        s, o = r.get("Subject",""), r.get("Object","")
        if s and o: G.add_edge(s.lower()[:40], o.lower()[:40])

    cycles = []
    try:
        cycles = list(nx.simple_cycles(G))
    except Exception:
        pass

    cycle_nodes = {n for c in cycles for n in c}

    for r in rels:
        s   = r.get("Subject","").lower()
        rel = r.get("Relationship","").lower()
        o   = r.get("Object","").lower()
        src = r.get("Source Sentence","")

        sev = None
        ptype = ""
        if any(w in rel for w in PROBLEM_WORDS) or any(w in s for w in PROBLEM_WORDS) \
                or any(w in o for w in PROBLEM_WORDS):
            sev = "High"; ptype = "problem keyword"
        elif s in cycle_nodes or o in cycle_nodes:
            sev = "Medium"; ptype = "circular dependency"
        else:
            ctx = src.lower()
            if any(w in ctx for w in PROBLEM_WORDS):
                sev = "Low"; ptype = "problematic context"

        if sev:
            problems.append({"Subject": r["Subject"], "Relationship": r["Relationship"],
                              "Object": r["Object"], "Problem Type": ptype,
                              "Severity": sev, "Source Sentence": src})
    return problems

# ── Knowledge graph ───────────────────────────────────────────────────────────
def build_knowledge_graph(rels: List[Dict], entities: List[Dict],
                          problems: List[Dict], max_nodes: int = 120) -> Dict:
    """Build a knowledge graph from extracted relationships + entities."""
    ent_type, ent_count = {}, {}
    for e in entities:
        key = e["Entity"].lower()
        if key not in ent_type:
            ent_type[key] = e["Type"]
            ent_count[key] = e["Occurrences"]

    problem_pairs = {(p["Subject"].lower()[:40], p["Object"].lower()[:40], )
                     for p in problems}

    G = nx.DiGraph()
    for r in rels:
        s = _clean(r.get("Subject", ""), 40)
        o = _clean(r.get("Object", ""), 40)
        if not s or not o or s.lower() == o.lower():
            continue
        sk, ok = s.lower(), o.lower()
        for node_key, label in ((sk, s), (ok, o)):
            if G.has_node(node_key):
                G.nodes[node_key]["weight"] += 1
            else:
                G.add_node(node_key, label=label,
                           etype=ent_type.get(node_key, ""),
                           weight=max(ent_count.get(node_key, 1), 1))
        if G.has_edge(sk, ok):
            G[sk][ok]["weight"] += 1
        else:
            G.add_edge(sk, ok, label=r.get("Relationship", ""),
                       rtype=r.get("Rel. Type", ""),
                       sentence=_clean(r.get("Source Sentence", ""), 200),
                       problem=(sk[:40], ok[:40]) in problem_pairs,
                       weight=1)

    # Keep the graph readable: most-connected nodes only, drop isolates
    if G.number_of_nodes() > max_nodes:
        keep = [n for n, _ in sorted(G.degree, key=lambda x: -x[1])[:max_nodes]]
        G = G.subgraph(keep).copy()
    G.remove_nodes_from(list(nx.isolates(G)))

    comm_of: Dict[str, int] = {}
    try:
        if G.number_of_nodes() > 2:
            communities = nx.community.greedy_modularity_communities(G.to_undirected())
            comm_of = {n: i for i, c in enumerate(communities) for n in c}
    except Exception:
        pass

    centrality = nx.degree_centrality(G) if G.number_of_nodes() else {}

    nodes = [{"id": n, "label": d["label"],
              "type": d.get("etype") or "Concept",
              "size": d.get("weight", 1),
              "centrality": round(centrality.get(n, 0.0), 4),
              "group": comm_of.get(n, 0)}
             for n, d in G.nodes(data=True)]
    edges = [{"from": u, "to": v,
              "label": d.get("label", ""), "rtype": d.get("rtype", ""),
              "sentence": d.get("sentence", ""),
              "problem": bool(d.get("problem")), "weight": d.get("weight", 1)}
             for u, v, d in G.edges(data=True)]
    return {"nodes": nodes, "edges": edges,
            "stats": {"nodes": len(nodes), "edges": len(edges),
                      "communities": len(set(comm_of.values())) if comm_of else 0}}

# ── Export helpers ────────────────────────────────────────────────────────────
def _export_excel_bytes(results: Dict) -> bytes:
    wb = Workbook()
    HF = Font(bold=True, color="FFFFFF", size=11)
    HB = PatternFill("solid", fgColor="1F4E79")
    RH = PatternFill("solid", fgColor="B71C1C")
    RF = PatternFill("solid", fgColor="FCE4EC")
    YF = PatternFill("solid", fgColor="FFF9C4")

    def ws(name, data, hfill=None):
        sheet = wb.create_sheet(name)
        if not data: sheet.append(["No data."]); return sheet
        df = pd.DataFrame(data)
        sheet.append(list(df.columns))
        for c in sheet[1]:
            c.font = HF; c.fill = hfill or HB
            c.alignment = Alignment(horizontal="center", wrap_text=True)
        for row in df.itertuples(index=False): sheet.append(list(row))
        for col in sheet.columns:
            w = max((len(str(c.value or "")) for c in col), default=10)
            sheet.column_dimensions[col[0].column_letter].width = min(w + 2, 55)
        return sheet

    s0 = wb.active; s0.title = "Summary"
    s0.append(["Metric", "Value"])
    for c in s0[1]: c.font = HF; c.fill = HB; c.alignment = Alignment(horizontal="center")
    for k, v in results["summary"].items(): s0.append([k, v])
    for col in s0.columns: s0.column_dimensions[col[0].column_letter].width = 35

    ws("Files",                  results["files"])
    ws("Entities",               results["entities"])
    ws("Key Concepts",           results["concepts"])
    ws("Variables",              results["variables"])
    ws("Relationships",          results["relationships"])
    sp = ws("Problem Relationships", results["problems"], hfill=RH)
    if results["problems"]:
        keys = list(results["problems"][0].keys())
        sev_col = keys.index("Severity") + 1
        for i in range(2, sp.max_row + 1):
            fill = RF if sp.cell(i, sev_col).value == "High" else YF
            for cell in sp[i]: cell.fill = fill

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def _export_csv_bytes(results: Dict) -> bytes:
    buf = io.BytesIO()
    sheets = {"summary": [results["summary"]], "files": results["files"],
              "entities": results["entities"], "concepts": results["concepts"],
              "variables": results["variables"], "relationships": results["relationships"],
              "problems": results["problems"]}
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in sheets.items():
            csv_str = pd.DataFrame(data if data else [{name: "No data"}]).to_csv(index=False, encoding="utf-8-sig")
            zf.writestr(f"{name}.csv", csv_str)
    return buf.getvalue()

def _export_markdown_bytes(results: Dict) -> bytes:
    def section(title, data):
        lines = [f"## {title}", ""]
        if not data: return lines + ["*No data.*", ""]
        df = pd.DataFrame(data)
        lines.append("| " + " | ".join(str(c) for c in df.columns) + " |")
        lines.append("|" + "|".join("---" for _ in df.columns) + "|")
        for _, row in df.iterrows():
            lines.append("| " + " | ".join(str(v)[:120].replace("|","\\|").replace("\n"," ") for v in row) + " |")
        return lines + [""]

    s = results["summary"]
    out = ["# Document Analysis Report", "", "## Summary", "",
           "| Metric | Value |", "|--------|-------|"]
    for k, v in s.items(): out.append(f"| {k} | {v} |")
    out += ["", "---", ""]
    out += section("Files Analyzed",        results["files"])
    out += section("Named Entities",        results["entities"])
    out += section("Key Concepts",          results["concepts"])
    out += section("Variables",             results["variables"])
    out += section("Relationships",         results["relationships"])
    out += section("Problem Relationships", results["problems"])
    return "\n".join(out).encode("utf-8")

# ── Flask app ─────────────────────────────────────────────────────────────────
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024  # 64 MB

_RESULTS: Dict[str, dict] = {}  # session_id → results

@app.route("/")
def index():
    return render_template("index.html", nlp_model=NLP_MODEL)

@app.route("/analyze", methods=["POST"])
def analyze():
    uploaded = request.files.getlist("files")
    if not uploaded or all(f.filename == "" for f in uploaded):
        return jsonify({"error": "No files uploaded"}), 400

    records, texts = [], []
    for f in uploaded:
        name = f.filename
        suffix = Path(name).suffix.lower() or ".tmp"
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp_path = tmp.name
                f.save(tmp_path)
            text, ftype = read_document(tmp_path)
            texts.append(text)
            records.append({"Filename": name, "Format": ftype,
                             "Words": len(text.split()), "Characters": len(text), "Status": "OK"})
        except Exception as exc:
            records.append({"Filename": name, "Format": "?",
                             "Words": 0, "Characters": 0, "Status": f"Error: {exc}"})
        finally:
            if tmp_path and os.path.exists(tmp_path):
                os.unlink(tmp_path)

    combined = "\n\n".join(texts)
    if not combined.strip():
        return jsonify({"error": "No text could be extracted from the uploaded files."}), 400
    if len(combined) > 400_000:
        combined = combined[:400_000]

    entities  = extract_entities(combined)
    concepts  = extract_concepts(combined)
    variables = extract_variables(combined)
    svo       = extract_svo(combined)
    causal    = extract_causal(combined)
    all_rels  = svo + causal
    problems  = identify_problems(all_rels, entities, combined)
    graph     = build_knowledge_graph(all_rels, entities, problems)

    sid = str(uuid.uuid4())
    results = {
        "files": records, "entities": entities, "concepts": concepts,
        "variables": variables, "relationships": all_rels, "problems": problems,
        "graph": graph,
        "summary": {
            "Files Analyzed":       len(records),
            "Total Words":          f"{sum(r.get('Words', 0) for r in records):,}",
            "Entities Found":       len(entities),
            "Key Concepts":         len(concepts),
            "Variables Detected":   len(variables),
            "Relationships Mapped": len(all_rels),
            "Problem Relationships":len(problems),
            "Graph Nodes":          graph["stats"]["nodes"],
            "Graph Edges":          graph["stats"]["edges"],
            "NLP Model":            NLP_MODEL,
            "Concept Model":        "all-MiniLM-L6-v2 (KeyBERT)",
        },
    }
    _RESULTS[sid] = results
    return jsonify({"session_id": sid, "results": results})

@app.route("/export/<sid>/<fmt>")
def export_results(sid: str, fmt: str):
    results = _RESULTS.get(sid)
    if not results:
        return "Session not found or expired.", 404

    if fmt == "xlsx":
        data = _export_excel_bytes(results)
        return send_file(io.BytesIO(data), download_name="document_analysis.xlsx",
                         as_attachment=True,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if fmt == "csv":
        data = _export_csv_bytes(results)
        return send_file(io.BytesIO(data), download_name="document_analysis.zip",
                         as_attachment=True, mimetype="application/zip")
    if fmt == "md":
        data = _export_markdown_bytes(results)
        return send_file(io.BytesIO(data), download_name="document_analysis.md",
                         as_attachment=True, mimetype="text/markdown")
    return "Unknown format.", 400

def _lan_ip() -> str:
    """Pick the real LAN address, preferring private ranges over VPN/CGNAT (100.64/10)."""
    import socket
    candidates = []
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ip = info[4][0]
            if not ip.startswith(("127.", "169.254.")):
                candidates.append(ip)
    except Exception:
        pass
    for ip in candidates:
        if ip.startswith(("192.168.", "10.")) or \
           (ip.startswith("172.") and 16 <= int(ip.split(".")[1]) <= 31):
            return ip
    return candidates[0] if candidates else "127.0.0.1"

if __name__ == "__main__":
    import webbrowser, threading
    lan = _lan_ip()
    url = "http://localhost:5000"
    threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    print(f"Data Analyzer Web running at:")
    print(f"  This computer : {url}")
    print(f"  On your WiFi  : http://{lan}:5000")
    app.run(host="0.0.0.0", port=5000, debug=False, threaded=True)

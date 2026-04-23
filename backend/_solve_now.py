import sys, re
sys.path.insert(0, "backend")
import pythoncom; pythoncom.CoInitialize()
import win32com.client as win32
from utils import find_matte_docx
from solver import solve_task, ensure_ollama_running
from pathlib import Path

MODEL = "qwen3:8b"
SEP   = "\u2500" * 44

doc_path = find_matte_docx()
word     = win32.GetActiveObject("Word.Application")
target   = Path(doc_path).name.lower()
doc = None
for i in range(1, word.Documents.Count + 1):
    d = word.Documents(i)
    if d.FullName.split("/")[-1].split("\\")[-1].lower() == target:
        doc = d
        break

print(f"Dok: {doc is not None}, paragrafer: {doc.Paragraphs.Count}", flush=True)

# Rydd feil-løsninger
i = 1
while i <= doc.Paragraphs.Count:
    t = doc.Paragraphs(i).Range.Text.strip()
    if t.startswith("\u2500") or t.startswith("\u2550"):
        if i + 1 <= doc.Paragraphs.Count:
            nxt = doc.Paragraphs(i + 1).Range.Text.strip().lower()
            if nxt.startswith("feil"):
                start = doc.Paragraphs(i).Range.Start
                j = i + 1
                while j <= doc.Paragraphs.Count:
                    txt = doc.Paragraphs(j).Range.Text.strip()
                    if j > i and (txt.startswith("\u2500") or txt.startswith("\u2550")):
                        break
                    j += 1
                end = doc.Paragraphs(min(j, doc.Paragraphs.Count)).Range.End
                doc.Range(start, end).Delete()
                print("Fjernet feil-losning", flush=True)
                continue
    i += 1

# Finn oppgaver
TRIG = re.compile(r"[-\u2013\u2014]\s*l[\u00f8o\u00f6]ss?[\s!.]*$", re.IGNORECASE)
tasks = []
for i in range(1, doc.Paragraphs.Count + 1):
    t = doc.Paragraphs(i).Range.Text.rstrip("\r\n\x07")
    if not TRIG.search(t):
        continue
    if i < doc.Paragraphs.Count:
        nxt = doc.Paragraphs(i + 1).Range.Text.strip()
        if nxt.startswith("\u2500") or nxt.startswith("\u2550"):
            nxt2 = ""
            if i + 2 <= doc.Paragraphs.Count:
                nxt2 = doc.Paragraphs(i + 2).Range.Text.strip().lower()
            if not nxt2.startswith("feil"):
                continue
    task_text = TRIG.sub("", t).strip()
    if task_text:
        tasks.append({"index": i, "text": task_text})

print(f"Fant {len(tasks)} oppgave(r)", flush=True)
if not tasks:
    sys.exit(0)

ensure_ollama_running(MODEL)

for task in sorted(tasks, key=lambda x: x["index"], reverse=True):
    print(f"Loser oppgave {task['index']}", flush=True)
    solution = solve_task(task["text"], MODEL)
    print(f"Svar (start): {solution[:100]}", flush=True)

    doc.AutoSaveOn = False
    lines = [SEP]
    for line in solution.strip().split("\n"):
        s = line.strip()
        if s:
            lines.append(s)
    lines.append(SEP)

    block = "\n" + "\n".join(lines)
    doc.Paragraphs(task["index"]).Range.InsertAfter(block)

    for line in lines:
        if line.lower().startswith("svar:"):
            f = doc.Content.Find
            f.ClearFormatting()
            f.Replacement.ClearFormatting()
            f.Text = line[:50]
            f.Replacement.Text = line[:50]
            f.Replacement.Font.Bold  = True
            f.Replacement.Font.Color = 0x006400
            f.Execute(Replace=1)

    doc.Save()
    doc.AutoSaveOn = True
    print("Lagret OK", flush=True)

print("Ferdig!")

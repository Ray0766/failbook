"""Static site: a case catalog and one evidence page per case, in English and Chinese.

Design notes (see docs/DECISIONS.md, "Site"): a document-style "research archive"
rather than a dashboard. Pages read top to bottom (task, expected, observed,
evidence, reproduce, retests), work without JavaScript, reflow to 320px, and
show verdicts as glyph + word, never colour alone. Counts, not rates.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import html
import json
import re
import shutil
import subprocess
from pathlib import Path

from . import __version__
from .case import Case, all_cases, repo_root
from .docs import CATEGORY_LABEL, KIND_LABEL, REPO_URL, case_results, failure_grades

T = {
    "en": {
        "site": "Failbook", "tagline": "Reproducible cases where an AI agent said “done” and the spreadsheet was wrong.",
        "cases": "Cases", "method": "How the checks are validated", "search": "Search cases, rules, tags",
        "showing": "Showing {n} of {total} cases", "caught_by": "Caught by", "difficulty": "Difficulty",
        "id": "ID", "failure": "Observable failure", "artifact": "Failure artifact", "evidence": "Evidence",
        "retests": "Retests", "no_runs": "no runs yet", "runs": "{n} run(s): {detail}",
        "rules_fail": "{f} of {n} rules fail", "rules_count": "{f} of {n} rules", "all": "All", "lang_other": "中文", "skip": "Skip to rule checks", "skip_main": "Skip to content",
        "about": "About this record", "about_text": "The failure artifact was produced by running <a href=\"{repo}/blob/main/cases/{dir}/failure.py\"><code>failure.py</code></a>, a script that makes the mistake described under “Observed in the failure artifact”. It reconstructs a commonly reported pattern; it is not a recording of a particular model.",
        "task": "Task", "task_note": "Given to the agent with the input file(s), canonical variant (seed 0).",
        "expected": "Expected", "observed": "Observed in the failure artifact", "claim": "Illustrative claim",
        "claim_note": "Written for this reconstruction, not quoted from a run.", "explanation": "Explanation",
        "rule_checks": "Rule checks", "legend": "Legend", "passed_rules": "Passed rules ({n})",
        "checks": "Checks", "human": "Human judgement", "human_none": "None recorded. Automatic rules only.",
        "reproduce": "Reproduce and retest", "copy": "Copy command", "copied": "Copied",
        "files": "Files", "file": "File", "role": "Role", "size": "Size",
        "role_input": "input given to the agent", "role_failure": "failure artifact", "role_grade": "grading of the failure artifact",
        "validation": "Is this checker trustworthy?", "fix": "Fix and self-check",
        "history": "Retest history", "history_note": "Append-only: every run is a new record. Counts, not rates; a curated case says nothing about overall reliability.",
        "date": "Date", "cversion": "Case version", "older": " (older)", "agent": "Agent", "model": "Model", "fixcol": "Fix", "status": "Run status", "verdict": "Verdict", "notpassed": "Rules not passed",
        "limitations": "Limitations", "classification": "Classification", "tool": "Tool behaviour involved",
        "prev": "Previous case", "next": "Next case", "built": "Built {date} from {commit} · failbook {ver}",
        "pass": "Pass", "fail": "Fail", "inconclusive": "Inconclusive", "notrun": "Not run",
        "grader": "Grader", "on_failure": "Failure artifact", "on_reference": "Correct solution", "on_noop": "Untouched input",
        "g_failbook": "Failbook (all rules, recalculated)", "g_values": "Values only (recalculated)", "g_cached": "Values only (stored values, no recalculation)",
        "right": "right verdict", "wrong": "wrong verdict",
        "selfcheck": "The candidate fix <a href=\"{repo}/tree/main/fixes/spreadsheet-verify-skill\">spreadsheet-verify-skill</a> asks the agent to run a task-agnostic audit before reporting success. On this case's failure artifact it reports:",
        "selfcheck_clean": "nothing. This failure leaves no generic symptom; only the task's requirements reveal it.",
        "fix_hypothesis": "Whether agents actually do better with the fix is not yet measured: no runs with and without it are recorded.",
        "empty_history": "No recorded runs yet. Prepare a fresh directory, give TASK.md to any agent, grade the result with <code>--record results</code>, and open a pull request with the new file.",
        "catalog_note": "A curated set of known failure patterns, not a random sample. Every failure artifact is a reconstruction until a recorded run reproduces it.",
        "how": "How to read a case", "how_text": "Each case page follows the evidence chain: the task and its input, what a correct result looks like, what the failure artifact contains, the rule-by-rule grading with cell-level evidence, how to reproduce it without any AI, and recorded retests.",
    },
    "zh": {
        "site": "Failbook", "tagline": "可复现的 AI 翻车案例：智能体说“完成了”，表格却是错的。",
        "cases": "案例", "method": "检查器如何被验证", "search": "搜索案例、规则、标签",
        "showing": "显示 {n} / {total} 个案例", "caught_by": "由哪类规则抓到", "difficulty": "难度",
        "id": "编号", "failure": "可观察到的失败", "artifact": "失败产物", "evidence": "证据级别",
        "retests": "复测", "no_runs": "尚无运行", "runs": "{n} 次运行：{detail}",
        "rules_fail": "{n} 条规则中 {f} 条不通过", "rules_count": "{f}/{n} 条规则", "all": "全部", "lang_other": "English", "skip": "跳到规则检查", "skip_main": "跳到正文",
        "about": "关于这条记录", "about_text": "失败产物由 <a href=\"{repo}/blob/main/cases/{dir}/failure.py\"><code>failure.py</code></a> 生成，它会犯下文“失败产物中观察到的情况”所述的错误。这是对常见失败模式的复现，不是某个模型的运行记录。",
        "task": "任务", "task_note": "连同输入文件一起交给智能体（标准变体，seed 0）。",
        "expected": "预期结果", "observed": "失败产物中观察到的情况", "claim": "示例声明",
        "claim_note": "为本复现撰写，并非摘自真实运行。", "explanation": "解释",
        "rule_checks": "规则检查", "legend": "图例", "passed_rules": "通过的规则（{n}）",
        "checks": "检查", "human": "人工判断", "human_none": "暂无记录，仅自动规则。",
        "reproduce": "复现与复测", "copy": "复制命令", "copied": "已复制",
        "files": "文件", "file": "文件", "role": "用途", "size": "大小",
        "role_input": "交给智能体的输入", "role_failure": "失败产物", "role_grade": "失败产物的判定结果",
        "validation": "这个检查器可信吗？", "fix": "修复与自检",
        "history": "复测记录", "history_note": "只追加：每次运行都是一条新记录。只显示计数，不算比率；精选案例不能说明总体可靠性。",
        "date": "日期", "cversion": "案例版本", "older": "（旧版）", "agent": "智能体", "model": "模型", "fixcol": "修复", "status": "运行状态", "verdict": "判定", "notpassed": "未通过的规则",
        "limitations": "限制", "classification": "分类", "tool": "涉及的工具行为",
        "prev": "上一个案例", "next": "下一个案例", "built": "构建于 {date}，提交 {commit} · failbook {ver}",
        "pass": "通过", "fail": "不通过", "inconclusive": "无法判定", "notrun": "未运行",
        "grader": "评分器", "on_failure": "失败产物", "on_reference": "正确解", "on_noop": "原样交回的输入",
        "g_failbook": "Failbook（全部规则，重新计算）", "g_values": "只查数值（重新计算）", "g_cached": "只查数值（读缓存值，不重算）",
        "right": "判定正确", "wrong": "判定错误",
        "selfcheck": "候选修复 <a href=\"{repo}/tree/main/fixes/spreadsheet-verify-skill\">spreadsheet-verify-skill</a> 要求智能体在汇报完成前运行一个与任务无关的审计脚本。它对本案例失败产物的报告：",
        "selfcheck_clean": "无。这个失败没有通用症状，只有对照任务要求才能发现。",
        "fix_hypothesis": "智能体使用该修复后是否真的做得更好，目前尚未测量：还没有加与不加修复的运行记录。",
        "empty_history": "尚无复测记录。准备一个新目录，把 TASK.md 交给任意智能体，用 <code>--record results</code> 评分，再把新文件提交 PR。",
        "catalog_note": "这是精选的已知失败模式，不是随机样本。在真实运行复现之前，每个失败产物都只是复现品。",
        "how": "如何阅读一个案例", "how_text": "每个案例页都沿着证据链展开：任务与输入、正确结果的样子、失败产物里有什么、带单元格级证据的逐条判定、不依赖任何 AI 的复现方法，以及复测记录。",
    },
}

GLYPH = {
    "pass": '<svg viewBox="0 0 16 16" aria-hidden="true"><circle cx="8" cy="8" r="6.5" fill="none" stroke="currentColor" stroke-width="1.5"/><path d="M5 8.2l2 2 4-4.4" fill="none" stroke="currentColor" stroke-width="1.6"/></svg>',
    "fail": '<svg viewBox="0 0 16 16" aria-hidden="true"><rect x="1.75" y="1.75" width="12.5" height="12.5" fill="none" stroke="currentColor" stroke-width="1.5"/><path d="M5.2 5.2l5.6 5.6M10.8 5.2l-5.6 5.6" stroke="currentColor" stroke-width="1.6"/></svg>',
    "inconclusive": '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M8 1.2L14.8 8 8 14.8 1.2 8z" fill="none" stroke="currentColor" stroke-width="1.5"/><text x="8" y="11" text-anchor="middle" font-size="8" fill="currentColor">?</text></svg>',
    "notrun": '<svg viewBox="0 0 16 16" aria-hidden="true"><circle cx="8" cy="8" r="6.5" fill="none" stroke="currentColor" stroke-width="1.5"/><path d="M5 8h6" stroke="currentColor" stroke-width="1.6"/></svg>',
}

CSS = r"""
:root{--bg:#FAFAF7;--surface:#FFFFFF;--sunken:#F1F0EB;--text:#1B1B18;--muted:#595851;--border:#D6D4CC;--border-strong:#85837B;
--link:#1D4ED8;--focus:#1D4ED8;--fail:#B3261E;--fail-tint:#FCEBE9;--pass:#1E6B3C;--pass-tint:#E6F2EA;--inc:#8A5300;--inc-tint:#FBF0DB;--neutral:#595851;--neutral-tint:#EDECE7;
--font-sans:system-ui,-apple-system,"Segoe UI",Roboto,"Helvetica Neue",Arial,"PingFang SC","Hiragino Sans GB","Microsoft YaHei","Noto Sans CJK SC","Source Han Sans SC","WenQuanYi Micro Hei",sans-serif;
--font-mono:ui-monospace,"SFMono-Regular","Cascadia Mono","Segoe UI Mono",Menlo,Consolas,"Liberation Mono","DejaVu Sans Mono","PingFang SC","Microsoft YaHei","Noto Sans Mono CJK SC",monospace;color-scheme:light dark}
@media (prefers-color-scheme:dark){:root{--bg:#141413;--surface:#1C1C1A;--sunken:#252522;--text:#ECEAE3;--muted:#A8A69E;--border:#3A3934;--border-strong:#7A786F;
--link:#8AB4F8;--focus:#8AB4F8;--fail:#F2938A;--fail-tint:#3A1E1B;--pass:#7FCF98;--pass-tint:#1B3223;--inc:#E9B95A;--inc-tint:#382C14;--neutral:#A8A69E;--neutral-tint:#2A2A27}}
*{box-sizing:border-box}html{scroll-padding-top:16px;-webkit-text-size-adjust:100%}
body{margin:0;background:var(--bg);color:var(--text);font:16px/1.6 var(--font-sans)}
:lang(zh-Hans) body,body:lang(zh-Hans){line-height:1.75}
a{color:var(--link)}a:hover{text-decoration-thickness:2px}
:focus-visible{outline:2px solid var(--focus);outline-offset:2px}
.skip{position:absolute;left:-999px;top:8px}.skip:focus{left:16px;background:var(--surface);padding:8px 12px;z-index:2}
.wrap{max-width:1120px;margin:0 auto;padding:0 16px}@media(min-width:640px){.wrap{padding:0 24px}}
header.site{border-bottom:1px solid var(--border);background:var(--surface)}
header.site .wrap{display:flex;flex-wrap:wrap;gap:8px 24px;align-items:baseline;padding-top:12px;padding-bottom:12px}
header.site .brand{font-weight:600;font-size:20px;text-decoration:none;color:var(--text)}
header.site nav{display:flex;flex-wrap:wrap;gap:4px 16px;font-size:14px}
main{padding:24px 0 64px}
h1{font-size:24px;line-height:1.3;margin:8px 0 8px;font-weight:600}@media(min-width:768px){h1{font-size:30px;line-height:1.25}}
h2{font-size:20px;line-height:1.35;margin:48px 0 12px;font-weight:600}
h3{font-size:16px;margin:24px 0 8px;font-weight:600}
p,li{max-width:72ch}
.muted{color:var(--muted)}.small{font-size:14px}.xs{font-size:13px}
code,pre,.mono{font-family:var(--font-mono);font-size:14px}
code{background:var(--sunken);padding:1px 4px;border-radius:3px;overflow-wrap:anywhere}
pre{background:var(--sunken);padding:12px 16px;overflow-x:auto;border-radius:4px;line-height:1.5;margin:0}
pre code{background:none;padding:0;overflow-wrap:normal}
.meta{display:flex;flex-wrap:wrap;gap:4px 16px;font-size:14px;color:var(--muted);margin:0 0 16px;padding:0;list-style:none}
.note{border-left:3px solid var(--border-strong);padding:8px 16px;background:var(--surface);font-size:14px;max-width:none}
.dossier{display:grid;gap:0;border-top:1px solid var(--border)}
.dossier>div{display:grid;grid-template-columns:minmax(0,1fr);gap:4px 24px;padding:16px 0;border-bottom:1px solid var(--border)}
@media(min-width:768px){.dossier>div{grid-template-columns:200px minmax(0,1fr)}}
.dossier dt{font-weight:600;font-size:14px}.dossier dd{margin:0;min-width:0}
.dossier .sub{font-size:13px;color:var(--muted);font-weight:400;display:block}
.task{background:var(--surface);border:1px solid var(--border);padding:16px;border-radius:4px}
.task table,.tbl{border-collapse:collapse;font-size:14px;margin:8px 0}
.task th,.task td,.tbl th,.tbl td{border-bottom:1px solid var(--border);padding:8px 12px;text-align:left;vertical-align:top}
.tbl th{font-weight:600;color:var(--muted);font-size:13px}
.tbl{width:100%;background:var(--surface)}
.num{font-variant-numeric:tabular-nums}
blockquote{margin:0;padding:0 0 0 12px;border-left:3px solid var(--border-strong)}
.verdict{display:inline-flex;align-items:center;gap:6px;font-weight:600;white-space:nowrap;font-size:14px}
.verdict svg{width:16px;height:16px;flex:none}
.verdict.fail{color:var(--fail)}.verdict.pass{color:var(--pass)}.verdict.inconclusive{color:var(--inc)}.verdict.notrun{color:var(--neutral)}
.summary-line{display:flex;flex-wrap:wrap;gap:8px 16px;align-items:center;padding:12px 16px;background:var(--fail-tint);border-left:3px solid var(--fail);margin:16px 0}
.rules{list-style:none;padding:0;margin:0;border-top:1px solid var(--border)}
.rules>li{max-width:none}
.rule{padding:12px 16px;border-bottom:1px solid var(--border);background:var(--surface)}
.rule.fail{background:var(--fail-tint);border-left:3px solid var(--fail)}
.rule.inconclusive{background:var(--inc-tint);border-left:3px dashed var(--inc)}
.rule .head{display:flex;flex-wrap:wrap;gap:4px 12px;align-items:baseline}
.rule .rid{font-family:var(--font-mono);font-size:13px;color:var(--muted)}
.rule .kind{font-size:13px;color:var(--muted);border:1px solid var(--border-strong);border-radius:10px;padding:0 8px}
.rule .title{font-weight:600}
.rule .req{font-size:14px;color:var(--muted);margin:4px 0 0}
.rule .detail{font-family:var(--font-mono);font-size:14px;margin:8px 0 0;overflow-wrap:anywhere}
.rule:target{outline:2px solid var(--focus);outline-offset:-2px}
details>summary{cursor:pointer;padding:12px 16px;background:var(--surface);border-bottom:1px solid var(--border);font-weight:600;font-size:14px}
.legend{display:flex;flex-wrap:wrap;gap:4px 16px;font-size:13px;color:var(--muted);margin:0 0 8px}
.cmd{position:relative;margin:8px 0}
.cmd button{position:absolute;top:6px;right:6px}
@media(min-width:640px){.cmd pre{padding-right:136px}}
@media(max-width:639px){.cmd button{position:static;margin-top:6px}}
button,.chip{font:inherit;font-size:14px;background:var(--surface);color:var(--text);border:1px solid var(--border-strong);border-radius:4px;padding:4px 10px;cursor:pointer}
.chip[aria-pressed="true"]{background:var(--text);color:var(--bg)}
.toolbar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin:16px 0 8px}
.toolbar input{font:inherit;font-size:16px;padding:6px 10px;border:1px solid var(--border-strong);border-radius:4px;background:var(--surface);color:var(--text);min-width:0;flex:1 1 240px}
.chips{display:flex;flex-wrap:wrap;gap:6px;align-items:center;font-size:14px}
.chips .label{color:var(--muted);margin-right:4px}
.catalog td .title{font-weight:600}
.catalog td.mono{white-space:nowrap}
.catalog a.case{text-decoration:none}.catalog a.case:hover .title{text-decoration:underline}
@media(max-width:639px){
 .catalog thead{position:absolute;left:-9999px}
 .catalog,.catalog tbody,.catalog tr,.catalog td{display:block;width:100%}
 .catalog tr{border-bottom:1px solid var(--border);padding:12px 0}
 .catalog td{border:0;padding:2px 0}
 .catalog td[data-label]::before{content:attr(data-label) ": ";color:var(--muted);font-size:13px}
 .tbl.stack thead{position:absolute;left:-9999px}
 .tbl.stack,.tbl.stack tbody,.tbl.stack tr,.tbl.stack td{display:block;width:100%}
 .tbl.stack tr{border-bottom:1px solid var(--border);padding:8px 0}.tbl.stack td{border:0;padding:2px 12px}
 .tbl.stack td[data-label]::before{content:attr(data-label) ": ";color:var(--muted);font-size:13px}
}
.pager{display:flex;justify-content:space-between;gap:16px;margin-top:48px;font-size:14px;flex-wrap:wrap}
footer{border-top:1px solid var(--border);padding:16px 0;font-size:13px;color:var(--muted)}
[hidden]{display:none!important}
@media (prefers-reduced-motion:reduce){*,*::before,*::after{transition:none!important;animation:none!important;scroll-behavior:auto!important}}
@media print{header.site nav,.toolbar,.cmd button,.skip{display:none}.rule,.dossier>div{break-inside:avoid}}
"""

JS = r"""
(function(){
  var copyLabel = document.documentElement.getAttribute('data-copied') || 'Copied';
  document.querySelectorAll('.cmd button').forEach(function(b){
    b.addEventListener('click', function(){
      var text = b.parentNode.querySelector('code').innerText;
      var done = function(){ var old=b.textContent; b.textContent=copyLabel; var live=document.getElementById('live'); if(live){live.textContent=copyLabel;} setTimeout(function(){b.textContent=old;},1500); };
      if (navigator.clipboard) { navigator.clipboard.writeText(text).then(done, function(){}); }
    });
  });
  var table = document.querySelector('table.catalog'); if(!table) return;
  var input = document.getElementById('q'), live = document.getElementById('live');
  var rows = Array.prototype.slice.call(table.querySelectorAll('tbody tr'));
  var chips = Array.prototype.slice.call(document.querySelectorAll('.chip[data-kind]'));
  var tmpl = live.getAttribute('data-template');
  function apply(){
    var q = (input.value||'').trim().toLowerCase();
    var kinds = chips.filter(function(c){return c.getAttribute('aria-pressed')==='true' && c.dataset.kind;}).map(function(c){return c.dataset.kind;});
    var n = 0;
    rows.forEach(function(r){
      var okq = !q || r.dataset.search.indexOf(q) >= 0;
      var okk = !kinds.length || kinds.every(function(k){return (' '+r.dataset.kinds+' ').indexOf(' '+k+' ')>=0;});
      r.hidden = !(okq && okk); if(!r.hidden) n++;
    });
    live.textContent = tmpl.replace('{n}', n).replace('{total}', rows.length);
    var params = new URLSearchParams(); if(q) params.set('q', q); if(kinds.length) params.set('kind', kinds.join(','));
    history.replaceState(null, '', params.toString() ? '?' + params.toString() : location.pathname);
  }
  chips.forEach(function(c){ c.addEventListener('click', function(){
    if(!c.dataset.kind){ chips.forEach(function(o){o.setAttribute('aria-pressed','false');}); c.setAttribute('aria-pressed','true'); }
    else { c.setAttribute('aria-pressed', c.getAttribute('aria-pressed')==='true' ? 'false' : 'true');
      var any = chips.some(function(o){return o.dataset.kind && o.getAttribute('aria-pressed')==='true';});
      chips.forEach(function(o){ if(!o.dataset.kind) o.setAttribute('aria-pressed', any ? 'false' : 'true'); }); }
    apply(); }); });
  input.addEventListener('input', apply);
  var p = new URLSearchParams(location.search);
  if(p.get('q')) input.value = p.get('q');
  if(p.get('kind')) p.get('kind').split(',').forEach(function(k){ chips.forEach(function(c){ if(c.dataset.kind===k) c.setAttribute('aria-pressed','true'); }); });
  if(p.get('q')||p.get('kind')){ chips.forEach(function(c){ if(!c.dataset.kind) c.setAttribute('aria-pressed','false'); }); }
  apply();
})();
"""


# -- small markdown renderer for task texts -----------------------------------------------------
def _inline(text: str) -> str:
    out = html.escape(text, quote=False)
    out = re.sub(r"`([^`]+)`", lambda m: f"<code>{m.group(1)}</code>", out)
    out = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", out)
    out = re.sub(r"(?<![*\w])\*([^*\s][^*]*?)\*(?![*\w])", r"<em>\1</em>", out)
    out = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", r'<a href="\2">\1</a>', out)
    return out


def markdown(text: str) -> str:
    lines = text.splitlines()
    out, i = [], 0
    while i < len(lines):
        line = lines[i]
        if line.startswith("```"):
            j = i + 1
            buf = []
            while j < len(lines) and not lines[j].startswith("```"):
                buf.append(lines[j])
                j += 1
            out.append("<pre><code>" + html.escape("\n".join(buf)) + "</code></pre>")
            i = j + 1
        elif line.startswith("    ") and line.strip():
            buf = []
            while i < len(lines) and (lines[i].startswith("    ") or not lines[i].strip()) and (lines[i].strip() or (i + 1 < len(lines) and lines[i + 1].startswith("    "))):
                buf.append(lines[i][4:])
                i += 1
            out.append("<pre><code>" + html.escape("\n".join(buf).rstrip()) + "</code></pre>")
        elif line.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].startswith("|"):
                rows.append([c.strip() for c in lines[i].strip().strip("|").split("|")])
                i += 1
            head, body = rows[0], [r for r in rows[1:] if not all(re.fullmatch(r":?-{3,}:?", c) for c in r)]
            t = '<div style="overflow-x:auto"><table><thead><tr>' + "".join(f"<th>{_inline(c)}</th>" for c in head) + "</tr></thead><tbody>"
            t += "".join("<tr>" + "".join(f'<td class="num">{_inline(c)}</td>' for c in r) + "</tr>" for r in body)
            out.append(t + "</tbody></table></div>")
        elif re.match(r"^\s*[-*] ", line):
            items = []
            while i < len(lines) and re.match(r"^\s*[-*] ", lines[i]):
                items.append(_inline(re.sub(r"^\s*[-*] ", "", lines[i])))
                i += 1
            out.append("<ul>" + "".join(f"<li>{x}</li>" for x in items) + "</ul>")
        elif line.strip() == "---":
            out.append("<hr>")
            i += 1
        elif re.match(r"^#{1,4} ", line):
            level = min(len(line.split(" ")[0]) + 1, 4)
            out.append(f"<h{level}>{_inline(line.split(' ', 1)[1])}</h{level}>")
            i += 1
        elif not line.strip():
            i += 1
        else:
            buf = []
            while i < len(lines) and lines[i].strip() and not re.match(r"^(\s*[-*] |\||```|    |#{1,4} )", lines[i]):
                buf.append(lines[i])
                i += 1
            out.append("<p>" + _inline(" ".join(buf)) + "</p>")
    return "\n".join(out)


# -- helpers ---------------------------------------------------------------------------------------
TERMS_ZH = {"reconstructed": "复现", "observed": "实测", "checker-verified": "评分器已验证", "draft": "草稿",
            "retired": "已退役", "easy": "易", "medium": "中", "hard": "难"}


def term(word: str, lang: str) -> str:
    return TERMS_ZH.get(word, word) if lang == "zh" else word


def join(items, lang: str) -> str:
    return ("、" if lang == "zh" else ", ").join(items)

def verdict(status: str, lang: str) -> str:
    key = status if status in GLYPH else "notrun"
    word = T[lang]["notrun" if key == "notrun" else key]
    return f'<span class="verdict {key}">{GLYPH[key]}{html.escape(word)}</span>'


def _commit() -> str:
    try:
        return subprocess.run(["git", "rev-parse", "--short", "HEAD"], capture_output=True, text=True,
                              cwd=repo_root(), timeout=10).stdout.strip() or "unknown"
    except Exception:
        return "unknown"


def _page(lang: str, title: str, body: str, *, depth: int, other_href: str, description: str = "",
          skip_target: str = "#main", live_template: str = "") -> str:
    t = T[lang]
    up = "../" * depth
    home = up + ("zh/" if lang == "zh" else "")
    html_lang = "zh-Hans" if lang == "zh" else "en"
    return f"""<!doctype html>
<html lang="{html_lang}" data-copied="{html.escape(t['copied'])}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{html.escape(title)}</title>
<meta name="description" content="{html.escape(description or t['tagline'])}">
<link rel="stylesheet" href="{up}assets/site.css">
</head>
<body>
<a class="skip" href="{skip_target}">{html.escape(t['skip'] if skip_target == '#rules' else t['skip_main'])}</a>
<header class="site"><div class="wrap">
<a class="brand" href="{home}index.html">Failbook</a>
<nav aria-label="Site">
<a href="{home}index.html">{t['cases']}</a>
<a href="{home}method.html">{t['method']}</a>
<a href="{REPO_URL}">GitHub</a>
<a href="{other_href}" hreflang="{'en' if lang == 'zh' else 'zh-Hans'}">{t['lang_other']}</a>
</nav></div></header>
<main id="main"><div class="wrap">
{body}
</div></main>
<footer><div class="wrap">{html.escape(t['built'].format(date=dt.date.today().isoformat(), commit=_commit(), ver=__version__))} · <a href="{REPO_URL}/blob/main/LICENSE">MIT</a> · synthetic data</div></footer>
<div id="live" class="skip" aria-live="polite" data-template="{html.escape(live_template)}"></div>
<script src="{up}assets/site.js" defer></script>
</body>
</html>
"""


def _ablation() -> dict:
    root = repo_root()
    out = {}
    for name in ("checker-ablation.json", "fix-self-check.json"):
        p = root / "docs" / "evidence" / name if root else None
        out[name] = json.loads(p.read_text(encoding="utf-8")) if p and p.is_file() else {"rows": []}
    return out


def _caught_kinds(case: Case) -> list[str]:
    kinds = {r.id: r.kind for r in case.rules}
    return sorted({kinds[g] for g, v in failure_grades(case).items() if v.get("status") == "fail" and g in kinds})


def _runs_summary(runs: list[dict], lang: str) -> str:
    if not runs:
        return T[lang]["no_runs"]
    counts = {}
    for r in runs:
        counts[r["verdict"]] = counts.get(r["verdict"], 0) + 1
    detail = ", ".join(f"{n} {T[lang][v]}" for v, n in sorted(counts.items()))
    return T[lang]["runs"].format(n=len(runs), detail=detail)


# -- pages -----------------------------------------------------------------------------------------
def catalog(cases: list[Case], lang: str) -> str:
    t = T[lang]
    kinds_all = ["content", "value", "live", "preserve", "structure"]
    counts = {k: sum(1 for c in cases if k in _caught_kinds(c)) for k in kinds_all}
    chips = [f'<button class="chip" type="button" aria-pressed="true">{t["all"]} ({len(cases)})</button>']
    chips += [f'<button class="chip" type="button" data-kind="{k}" aria-pressed="false">{KIND_LABEL[lang][k]} ({counts[k]})</button>'
              for k in kinds_all if counts[k]]
    rows = []
    prefix = ""
    for c in cases:
        grades = failure_grades(c)
        nfail = sum(1 for g in grades.values() if g.get("status") == "fail")
        kinds = _caught_kinds(c)
        title = c.title_zh if lang == "zh" else c.title
        search = " ".join([c.id, c.title, c.title_zh, " ".join(c.spec.get("tags", [])), " ".join(r.id for r in c.rules)]).lower()
        runs = case_results(c)
        rows.append(
            f'<tr data-search="{html.escape(search)}" data-kinds="{" ".join(kinds)}">'
            f'<td class="mono" data-label="{t["id"]}">{c.id}</td>'
            f'<td><a class="case" href="{prefix}cases/{c.id}/index.html"><span class="title">{html.escape(title)}</span></a>'
            f'<div class="xs muted">{" · ".join(html.escape(x) for x in c.spec.get("tags", [])[:4])}</div></td>'
            f'<td data-label="{t["caught_by"]}">{join([KIND_LABEL[lang][k] for k in kinds], lang)}</td>'
            f'<td data-label="{t["artifact"]}">{verdict("fail" if nfail else "pass", lang)} '
            f'<span class="small num">{t["rules_count"].format(f=nfail, n=len(c.rules))}</span></td>'
            f'<td data-label="{t["evidence"]}" class="small">{html.escape(term(c.spec.get("classification", {}).get("evidence", "?"), lang))}</td>'
            f'<td data-label="{t["retests"]}" class="small">{html.escape(_runs_summary(runs, lang))}</td>'
            f"</tr>"
        )
    body = f"""
<h1>{html.escape(t['tagline'])}</h1>
<p class="muted">{html.escape(t['catalog_note'])}</p>
<div class="toolbar" role="search">
<label class="skip" for="q">{t['search']}</label>
<input id="q" type="search" placeholder="{html.escape(t['search'])}" autocomplete="off">
</div>
<div class="chips" role="group" aria-label="{t['caught_by']}"><span class="label">{t['caught_by']}:</span>{''.join(chips)}</div>
<p class="small muted" id="count" aria-live="polite">{t['showing'].format(n=len(cases), total=len(cases))}</p>
<table class="tbl catalog">
<thead><tr><th scope="col">{t['id']}</th><th scope="col">{t['failure']}</th><th scope="col">{t['caught_by']}</th><th scope="col">{t['artifact']}</th><th scope="col">{t['evidence']}</th><th scope="col">{t['retests']}</th></tr></thead>
<tbody>
{''.join(rows)}
</tbody></table>
<h2>{t['how']}</h2>
<p>{html.escape(t['how_text'])}</p>
"""
    return _page(lang, "Failbook" if lang == "en" else "Failbook · 中文", body, depth=1 if lang == "zh" else 0,
                 other_href=("../index.html" if lang == "zh" else "zh/index.html"), live_template=t["showing"])


def case_page(case: Case, cases: list[Case], lang: str, ablation: dict) -> str:
    t = T[lang]
    meta = case.params(0)
    grades = failure_grades(case)
    cls = case.spec.get("classification", {})
    title = case.title_zh if lang == "zh" else case.title
    nfail = sum(1 for g in grades.values() if g.get("status") == "fail")
    idx = [c.id for c in cases].index(case.id)
    prev_c = cases[idx - 1] if idx > 0 else None
    next_c = cases[idx + 1] if idx + 1 < len(cases) else None

    def rule_li(r, g):
        status = g.get("status", "notrun")
        rtitle = r.title_zh if lang == "zh" and r.title_zh else r.title
        detail = f'<p class="detail">{html.escape(g["detail"])}</p>' if g.get("detail") else ""
        return (f'<li class="rule {status}" id="rule-{r.id}"><div class="head">{verdict(status, lang)}'
                f'<span class="title">{html.escape(rtitle)}</span><span class="rid">{r.id}</span>'
                f'<span class="kind">{KIND_LABEL[lang][r.kind]}</span></div>'
                f'<p class="req">{t["checks"]}: “{html.escape(r.requirement)}”</p>{detail}</li>')

    failing = [rule_li(r, grades.get(r.id, {})) for r in case.rules if grades.get(r.id, {}).get("status") != "pass"]
    passing = [rule_li(r, grades.get(r.id, {})) for r in case.rules if grades.get(r.id, {}).get("status") == "pass"]

    cmds = [
        f"failbook prepare {case.id} /tmp/{case.id}\nfailbook solve {case.id} /tmp/{case.id} --kind failure\nfailbook grade {case.id} /tmp/{case.id}",
        f"failbook prepare {case.id} /tmp/{case.id}-try\n# give /tmp/{case.id}-try/TASK.md to an agent, then:\nfailbook grade {case.id} /tmp/{case.id}-try --record results --agent <agent> --model <model>",
        f"failbook export harbor {case.id} --out dist/harbor\nharbor run -p dist/harbor -a claude-code -m <model> -k 3",
    ]
    labels = (["Reproduce the failure (no AI, needs LibreOffice)", "Retest any agent", "Run in a Docker sandbox with Harbor"] if lang == "en"
              else ["复现失败（不需要 AI，需要 LibreOffice）", "复测任意智能体", "用 Harbor 在 Docker 沙箱中运行"])
    cmd_html = "".join(
        f'<h3>{html.escape(lab)}</h3><div class="cmd"><pre tabindex="0"><code>{html.escape(cmd)}</code></pre>'
        f'<button type="button">{t["copy"]}</button></div>' for lab, cmd in zip(labels, cmds))

    files = []
    for name in case.inputs:
        files.append((f"files/input/{name}", case.input_dir / name, t["role_input"]))
    for name in case.outputs:
        files.append((f"files/failure/{name}", case.failure_dir / name, t["role_failure"]))
    files.append(("files/failure/grade.json", case.failure_dir / "grade.json", t["role_grade"]))
    file_rows = "".join(
        f'<tr><td data-label="{t["file"]}"><a href="{href}" download>{html.escape(href.split("files/")[1])}</a></td>'
        f'<td data-label="{t["role"]}">{html.escape(role)}</td><td data-label="{t["size"]}" class="num">{p.stat().st_size:,} B</td>'
        f'<td data-label="SHA-256" class="mono xs">{hashlib.sha256(p.read_bytes()).hexdigest()[:16]}…</td></tr>'
        for href, p, role in files if p.is_file())

    # Validation: ablation rows for this case.
    arows = [r for r in ablation["checker-ablation.json"]["rows"] if r["case"] == case.id]
    def acell(output, grader):
        r = next((x for x in arows if x["output"] == output and x["grader"] == grader), None)
        if r is None:
            return "—"
        ok = r["verdict"] == r["expected"]
        word = t["right"] if ok else t["wrong"]
        return f'{verdict(r["verdict"], lang)} <span class="xs muted">({word})</span>'
    graders = [("failbook", t["g_failbook"]), ("values-only", t["g_values"]), ("cached-values", t["g_cached"])]
    val_rows = "".join(
        f'<tr><th scope="row">{html.escape(label)}</th>'
        + "".join(f'<td data-label="{t[col]}">{acell(out, g)}</td>' for out, col in (("failure", "on_failure"), ("reference", "on_reference"), ("noop", "on_noop")))
        + "</tr>" for g, label in graders)
    srow = next((r for r in ablation["fix-self-check.json"]["rows"] if r["case"] == case.id and r["output"] == "failure"), None)
    if srow and (srow["errors"] or srow["warnings"]):
        found = "<ul>" + "".join(
            f'<li><span class="mono xs">[{f["level"]}] {html.escape(f["check"])}</span> {html.escape(f["where"])}: {html.escape(f["message"])}</li>'
            for f in srow["errors"] + srow["warnings"]) + "</ul>"
    else:
        found = f"<p>{html.escape(t['selfcheck_clean'])}</p>"

    runs = case_results(case)
    if runs:
        hist = "".join(
            f'<tr><td data-label="{t["date"]}" class="num">{r["started_at"][:10]}</td>'
            f'<td data-label="{t["cversion"]}" class="num">v{r["case_version"]}{"" if r["case_version"] == case.version else t["older"]}</td>'
            f'<td data-label="{t["agent"]}">{html.escape(r["agent"]["name"] + " " + (r["agent"].get("version") or ""))}</td>'
            f'<td data-label="{t["model"]}">{html.escape(r.get("model") or "—")}</td><td data-label="{t["fixcol"]}">{html.escape(r.get("fix") or "—")}</td>'
            f'<td data-label="{t["status"]}">{html.escape(r.get("run_status", "—"))}</td><td data-label="{t["verdict"]}">{verdict(r["verdict"], lang)}</td>'
            f'<td data-label="{t["notpassed"]}" class="mono xs">{html.escape(", ".join(x["id"] for x in r["rules"] if x["status"] != "pass") or "—")}</td></tr>'
            for r in runs)
        history = (f'<table class="tbl stack"><thead><tr><th>{t["date"]}</th><th>{t["cversion"]}</th><th>{t["agent"]}</th><th>{t["model"]}</th><th>{t["fixcol"]}</th>'
                   f'<th>{t["status"]}</th><th>{t["verdict"]}</th><th>{t["notpassed"]}</th></tr></thead><tbody>{hist}</tbody></table>')
    else:
        history = f'<p>{t["empty_history"]}</p>'

    limits = "".join(f"<li>{html.escape(item.get(lang, item['en']))}</li>" for item in case.spec.get("limitations", []))
    task_html = markdown(case.task(meta, lang))
    claim = html.escape(case.text("failure.claim", lang))
    up = "../../" + ("../" if lang == "zh" else "")
    other = (f"{up}cases/{case.id}/index.html" if lang == "zh" else f"../../zh/cases/{case.id}/index.html")
    pager = '<nav class="pager" aria-label="Cases">' + (
        f'<a href="../{prev_c.id}/index.html">← {t["prev"]}: {prev_c.id}</a>' if prev_c else "<span></span>") + (
        f'<a href="../{next_c.id}/index.html">{t["next"]}: {next_c.id} →</a>' if next_c else "<span></span>") + "</nav>"

    body = f"""
<p class="small"><a href="{'../../' }index.html">{t['cases']}</a> / <span class="mono">{case.id}</span></p>
<h1>{html.escape(title)}</h1>
<ul class="meta"><li class="mono">{case.id}</li><li>{t['difficulty']}: {html.escape(term(case.spec.get('difficulty', ''), lang))}</li><li>{'案例版本' if lang == 'zh' else 'case version'} {case.version}</li>
<li>{html.escape(term(str(case.spec.get('status', '')).split()[0], lang))}</li><li>{t['evidence']}: {html.escape(term(cls.get('evidence', '?'), lang))}</li>
<li>{html.escape(CATEGORY_LABEL[lang].get(cls.get('category', 'undetermined'), ''))}</li></ul>
<div class="summary-line">{verdict('fail' if nfail else 'pass', lang)}<span>{t['artifact']}: {t['rules_fail'].format(f=nfail, n=len(case.rules))}</span>
<a class="small" href="#rules">{t['rule_checks']} ↓</a></div>
<p class="note"><strong>{t['about']}.</strong> {t['about_text'].format(repo=REPO_URL, dir=case.dir.name)}</p>

<dl class="dossier">
<div><dt>{t['task']}<span class="sub">{t['task_note']}</span></dt><dd><div class="task">{task_html}</div></dd></div>
<div><dt>{t['expected']}</dt><dd><p>{html.escape(case.text('failure.expected', lang))}</p></dd></div>
<div><dt>{t['observed']}</dt><dd><p>{html.escape(case.text('failure.observed', lang))}</p></dd></div>
<div><dt>{t['claim']}<span class="sub">{t['claim_note']}</span></dt><dd><blockquote><p>{claim}</p></blockquote></dd></div>
<div><dt>{t['explanation']}</dt><dd><p>{html.escape(case.text('failure.explanation', lang))}</p>
<p class="small muted">{t['classification']}: {html.escape(CATEGORY_LABEL[lang].get(cls.get('category', 'undetermined'), ''))} ({html.escape(term(cls.get('evidence', '?'), lang))}).
{t['tool']}: {html.escape(case.text('classification.tool_behaviour'))}</p></dd></div>
</dl>

<h2 id="rules">{t['rule_checks']} ({len(case.rules)})</h2>
<p class="legend"><span>{t['legend']}:</span>{verdict('fail', lang)}{verdict('pass', lang)}{verdict('inconclusive', lang)}{verdict('notrun', lang)}</p>
<ol class="rules" aria-label="{t['rule_checks']}">{''.join(failing)}</ol>
<details><summary>{t['passed_rules'].format(n=len(passing))}</summary><ol class="rules">{''.join(passing)}</ol></details>
<h3>{t['human']} (0)</h3><p class="small muted">{t['human_none']}</p>

<h2>{t['validation']}</h2>
<table class="tbl stack"><thead><tr><th scope="col">{t['grader']}</th><th scope="col">{t['on_failure']}</th><th scope="col">{t['on_reference']}</th><th scope="col">{t['on_noop']}</th></tr></thead>
<tbody>{val_rows}</tbody></table>
<p class="small muted"><a href="{'../../'}method.html">{t['method']}</a></p>

<h2>{t['fix']}</h2>
<p>{t['selfcheck'].format(repo=REPO_URL)}</p>
{found}
<p class="small muted">{html.escape(t['fix_hypothesis'])}</p>

<h2>{t['reproduce']}</h2>
{cmd_html}

<h2>{t['files']} ({len(files)})</h2>
<table class="tbl stack"><thead><tr><th>{t['file']}</th><th>{t['role']}</th><th>{t['size']}</th><th>SHA-256</th></tr></thead><tbody>{file_rows}</tbody></table>

<h2>{t['history']} ({len(runs)})</h2>
<p class="small muted">{html.escape(t['history_note'])}</p>
{history}

<h2>{t['limitations']}</h2>
<ul>{limits}</ul>
{pager}
"""
    return _page(lang, f"{case.id} · {title} · Failbook", body, depth=2 + (1 if lang == "zh" else 0), other_href=other,
                 description=case.text("failure.observed", lang)[:200], skip_target="#rules")


def method_page(cases: list[Case], lang: str) -> str:
    root = repo_root()
    parts = []
    for name in ("checker-ablation.md", "fix-self-check.md"):
        p = root / "docs" / "evidence" / name
        if p.is_file():
            text = p.read_text(encoding="utf-8")
            text = re.sub(r"<!--.*?-->", "", text)
            text = re.sub(r"\]\(\.\./\.\./cases/(xlsx-\d+)[^)]*\)", r"](cases/\1/index.html)", text)
            text = text.replace("](../../fixes/", f"]({REPO_URL}/tree/main/fixes/")
            parts.append(text)
    intro = ("<p>Every case's checker is tested before it is trusted: a correct reference solution (and alternative "
             "solutions) must pass, the failure artifact must fail on the rules the case names, the untouched input must "
             "fail, and the reference must still pass after LibreOffice re-saves it. CI runs this on three data variants "
             "(<code>failbook verify --seeds 3</code>). The two studies below are generated from the repository and can be "
             "reproduced with the commands they name.</p>" if lang == "en" else
             "<p>每个案例的检查器在被信任前都要经过检验：正确的参考解（及替代解法）必须通过；失败产物必须在案例列出的规则上失败；"
             "原样交回的输入必须失败；参考解经 LibreOffice 重新保存后仍须通过。CI 在三个数据变体上运行这些检验"
             "（<code>failbook verify --seeds 3</code>）。以下两项研究由仓库内容生成，可按其中给出的命令复现（正文为英文）。</p>")
    body = "<h1>" + html.escape(T[lang]["method"]) + "</h1>" + intro + "".join(
        markdown(re.sub(r"^# .*\n", "", p, count=1)).replace("<h2>", "<h2>") for p in parts)
    depth = 1 if lang == "zh" else 0
    return _page(lang, T[lang]["method"] + " · Failbook", body, depth=depth,
                 other_href=("../method.html" if lang == "zh" else "zh/method.html"))


def build_site(out: Path, results_dir: Path | None = None) -> Path:
    out = Path(out)
    if out.exists():
        shutil.rmtree(out)
    (out / "assets").mkdir(parents=True)
    (out / "assets" / "site.css").write_text(CSS.strip() + "\n", encoding="utf-8")
    (out / "assets" / "site.js").write_text(JS.strip() + "\n", encoding="utf-8")
    cases = all_cases()
    ablation = _ablation()
    for lang in ("en", "zh"):
        base = out if lang == "en" else out / "zh"
        base.mkdir(parents=True, exist_ok=True)
        (base / "index.html").write_text(catalog(cases, lang), encoding="utf-8")
        (base / "method.html").write_text(method_page(cases, lang), encoding="utf-8")
        for c in cases:
            d = base / "cases" / c.id
            d.mkdir(parents=True)
            (d / "index.html").write_text(case_page(c, cases, lang, ablation), encoding="utf-8")
            # Copies, not symlinks: GitHub Pages artifacts do not reliably keep links.
            shutil.copytree(c.input_dir, d / "files" / "input")
            shutil.copytree(c.failure_dir, d / "files" / "failure")
    (out / ".nojekyll").write_text("", encoding="utf-8")
    return out

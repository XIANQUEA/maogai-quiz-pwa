# 毛概期末刷题 PWA Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 构建一款可安装、可离线、无需账号的手机端毛概刷题 PWA，支持章节练习、随机练习、错题本、即时判题和本地备份恢复。

**Architecture:** 应用使用原生 ES Modules，将题库、判分、练习会话、学习记录、IndexedDB、备份和界面渲染分成独立模块。题库由一次性 Word 提取工具转换为版本化 JSON；浏览器端只读取已校验的数据。Service Worker 缓存静态资源与题库，个人记录始终保存在 IndexedDB 中。

**Tech Stack:** HTML5、CSS3、原生 JavaScript ES Modules、IndexedDB、Web App Manifest、Service Worker、Node.js 内置 `node:test`、PowerShell Word COM（仅用于一次性提取旧版 `.doc`）

---

## 0. 目标文件结构

```text
maogai-quiz-pwa/
├─ index.html                         # PWA HTML 外壳
├─ manifest.webmanifest               # 安装名称、图标和启动参数
├─ service-worker.js                  # 版本化离线缓存
├─ package.json                       # 本地服务与测试命令
├─ assets/
│  ├─ icons/icon-192.png
│  └─ icons/icon-512.png
├─ data/
│  ├─ questions.v1.json               # 正式题库
│  └─ question-bank-report.json       # 导入数量及异常报告
├─ src/
│  ├─ app.js                          # 启动、路由和模块协调
│  ├─ config.js                       # 应用和题库版本常量
│  ├─ domain/
│  │  ├─ questions.js                 # 题目结构校验与判分
│  │  ├─ session.js                   # 章节、随机、错题会话
│  │  └─ progress.js                  # 作答记录、错题规则和统计
│  ├─ data/
│  │  ├─ question-repository.js       # 题库加载
│  │  ├─ progress-store.js            # IndexedDB 持久化
│  │  └─ backup.js                    # 备份编码、校验和恢复
│  └─ ui/
│     ├─ screens.js                   # 各页面纯 HTML 渲染函数
│     ├─ controller.js                # 点击、选择和提交事件
│     └─ install.js                   # PWA 安装提示
├─ styles/
│  ├─ tokens.css                      # 颜色、字号、间距和圆角
│  ├─ base.css                        # 重置、排版和无障碍基础
│  └─ components.css                  # 卡片、选项、导航和反馈
├─ tools/
│  ├─ extract-doc.ps1                 # 从旧版 Word 文档提取纯文本
│  ├─ build-question-bank.mjs         # 文本转结构化题库
│  └─ serve.mjs                       # 零依赖静态开发服务器
├─ source/
│  └─ .gitkeep                        # 本地原始文档/文本不提交
├─ test/
│  ├─ fixtures/question-bank-sample.txt
│  ├─ questions.test.mjs
│  ├─ session.test.mjs
│  ├─ progress.test.mjs
│  ├─ backup.test.mjs
│  └─ bank-builder.test.mjs
└─ docs/superpowers/
   ├─ specs/2026-07-02-maogai-quiz-pwa-design.md
   └─ plans/2026-07-02-maogai-quiz-pwa-implementation.md
```

`source/` 中的原始 `.doc` 和提取文本只用于本地构建，不进入 Git。正式交付使用 `data/questions.v1.json`。

### Task 1: 建立零依赖项目骨架和本地服务器

**Files:**
- Create: `package.json`
- Create: `.gitignore`
- Create: `index.html`
- Create: `src/config.js`
- Create: `tools/serve.mjs`
- Create: `source/.gitkeep`

- [ ] **Step 1: 写入项目命令和忽略规则**

`package.json`：

```json
{
  "name": "maogai-quiz-pwa",
  "version": "1.0.0",
  "private": true,
  "type": "module",
  "scripts": {
    "serve": "node tools/serve.mjs",
    "test": "node --test test/*.test.mjs",
    "bank:build": "node tools/build-question-bank.mjs source/question-bank.txt data/questions.v1.json data/question-bank-report.json"
  }
}
```

`.gitignore`：

```gitignore
source/*
!source/.gitkeep
.superpowers/
*.log
```

- [ ] **Step 2: 创建最小 HTML 外壳和版本常量**

`index.html`：

```html
<!doctype html>
<html lang="zh-CN">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
    <meta name="theme-color" content="#2563eb">
    <meta name="description" content="可离线使用的毛概期末刷题应用">
    <link rel="manifest" href="./manifest.webmanifest">
    <link rel="stylesheet" href="./styles/tokens.css">
    <link rel="stylesheet" href="./styles/base.css">
    <link rel="stylesheet" href="./styles/components.css">
    <title>毛概刷题</title>
  </head>
  <body>
    <div id="app" aria-live="polite">
      <p class="app-loading">正在加载题库…</p>
    </div>
    <script type="module" src="./src/app.js"></script>
  </body>
</html>
```

`src/config.js`：

```js
export const APP_VERSION = "1.0.0";
export const BANK_VERSION = 1;
export const BANK_URL = "./data/questions.v1.json";
export const DB_NAME = "maogai-quiz";
export const DB_VERSION = 1;
export const BACKUP_FORMAT = "maogai-quiz-progress";
export const BACKUP_VERSION = 1;
```

- [ ] **Step 3: 实现零依赖静态服务器**

`tools/serve.mjs`：

```js
import { createServer } from "node:http";
import { createReadStream, statSync } from "node:fs";
import { extname, join, normalize } from "node:path";

const root = process.cwd();
const port = Number(process.env.PORT || 4173);
const types = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".webmanifest": "application/manifest+json; charset=utf-8",
  ".png": "image/png"
};

createServer((request, response) => {
  const pathname = new URL(request.url, `http://${request.headers.host}`).pathname;
  const relative = pathname === "/" ? "index.html" : decodeURIComponent(pathname.slice(1));
  const file = normalize(join(root, relative));
  if (!file.startsWith(root)) {
    response.writeHead(403).end("Forbidden");
    return;
  }
  try {
    if (!statSync(file).isFile()) throw new Error("not a file");
    response.writeHead(200, {
      "Content-Type": types[extname(file)] || "application/octet-stream",
      "Cache-Control": "no-cache"
    });
    createReadStream(file).pipe(response);
  } catch {
    response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    response.end("Not found");
  }
}).listen(port, "127.0.0.1", () => {
  console.log(`毛概刷题开发服务器：http://127.0.0.1:${port}`);
});
```

- [ ] **Step 4: 验证服务器**

Run: `npm run serve`

Expected: 输出 `毛概刷题开发服务器：http://127.0.0.1:4173`，访问后显示“正在加载题库…”且无 404。

- [ ] **Step 5: 提交骨架**

```bash
git add package.json .gitignore index.html src/config.js tools/serve.mjs source/.gitkeep
git commit -m "chore: scaffold static quiz app"
```

### Task 2: 题目校验与确定性判分

**Files:**
- Create: `src/domain/questions.js`
- Create: `test/questions.test.mjs`

- [ ] **Step 1: 写失败测试**

`test/questions.test.mjs`：

```js
import test from "node:test";
import assert from "node:assert/strict";
import {
  assertQuestion,
  gradeQuestion,
  normalizeSelection
} from "../src/domain/questions.js";

const single = {
  id: "chapter-00-single-004",
  chapterId: "chapter-00",
  chapterTitle: "导论",
  type: "single",
  sourceNumber: 4,
  stem: "最新理论成果是？",
  options: [
    { key: "A", text: "毛泽东思想" },
    { key: "C", text: "习近平新时代中国特色社会主义思想" }
  ],
  answer: ["C"],
  bankVersion: 1
};

test("单选题答案一致时正确", () => {
  assert.equal(gradeQuestion(single, ["C"]), true);
  assert.equal(gradeQuestion(single, ["A"]), false);
});

test("多选题忽略选择顺序但不允许少选", () => {
  const multiple = { ...single, type: "multiple", answer: ["A", "C"] };
  assert.equal(gradeQuestion(multiple, ["C", "A"]), true);
  assert.equal(gradeQuestion(multiple, ["A"]), false);
});

test("判断题只接受 T 或 F", () => {
  const judgment = {
    ...single,
    type: "judgment",
    options: [],
    answer: ["T"]
  };
  assert.equal(gradeQuestion(judgment, ["T"]), true);
  assert.throws(() => assertQuestion({ ...judgment, answer: ["A"] }), /判断题答案/);
});

test("选择值会去重、排序并转大写", () => {
  assert.deepEqual(normalizeSelection(["c", "A", "c"]), ["A", "C"]);
});

test("答案引用不存在选项时拒绝题目", () => {
  assert.throws(() => assertQuestion({ ...single, answer: ["D"] }), /不存在选项/);
});
```

- [ ] **Step 2: 运行并确认失败**

Run: `node --test test/questions.test.mjs`

Expected: FAIL，错误包含 `ERR_MODULE_NOT_FOUND`。

- [ ] **Step 3: 实现题目校验与判分**

`src/domain/questions.js`：

```js
const TYPES = new Set(["judgment", "single", "multiple"]);

export function normalizeSelection(values) {
  return [...new Set(values.map(value => String(value).trim().toUpperCase()))].sort();
}

export function assertQuestion(question) {
  if (!question || typeof question !== "object") throw new Error("题目必须是对象");
  for (const key of ["id", "chapterId", "chapterTitle", "type", "stem"]) {
    if (!String(question[key] ?? "").trim()) throw new Error(`题目缺少 ${key}`);
  }
  if (!TYPES.has(question.type)) throw new Error(`未知题型 ${question.type}`);
  if (!Array.isArray(question.answer) || question.answer.length === 0) {
    throw new Error("题目缺少答案");
  }
  const answer = normalizeSelection(question.answer);
  if (question.type === "judgment") {
    if (answer.length !== 1 || !["T", "F"].includes(answer[0])) {
      throw new Error("判断题答案必须是 T 或 F");
    }
    return true;
  }
  const options = Array.isArray(question.options) ? question.options : [];
  const keys = new Set(options.map(option => option.key));
  if (keys.size !== options.length || options.length < 2) throw new Error("选择题选项无效");
  for (const value of answer) {
    if (!keys.has(value)) throw new Error(`答案 ${value} 不存在选项`);
  }
  if (question.type === "single" && answer.length !== 1) {
    throw new Error("单选题只能有一个答案");
  }
  return true;
}

export function gradeQuestion(question, selected) {
  assertQuestion(question);
  const actual = normalizeSelection(selected);
  const expected = normalizeSelection(question.answer);
  return actual.length === expected.length &&
    actual.every((value, index) => value === expected[index]);
}
```

- [ ] **Step 4: 运行测试**

Run: `node --test test/questions.test.mjs`

Expected: 5 tests PASS。

- [ ] **Step 5: 提交**

```bash
git add src/domain/questions.js test/questions.test.mjs
git commit -m "feat: add question validation and grading"
```

### Task 3: 从旧版 Word 文档构建题库

**Files:**
- Create: `tools/extract-doc.ps1`
- Create: `tools/build-question-bank.mjs`
- Create: `test/fixtures/question-bank-sample.txt`
- Create: `test/bank-builder.test.mjs`
- Create: `data/questions.v1.json`
- Create: `data/question-bank-report.json`

- [ ] **Step 1: 创建可重复的 Word 文本提取脚本**

`tools/extract-doc.ps1`：

```powershell
param(
  [Parameter(Mandatory=$true)][string]$InputDoc,
  [Parameter(Mandatory=$true)][string]$OutputText
)
$word = $null
$doc = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Open((Resolve-Path $InputDoc), $false, $true)
  $text = $doc.Content.Text -replace "`v", "" -replace "`f", "`r"
  [System.IO.File]::WriteAllText(
    [System.IO.Path]::GetFullPath($OutputText),
    $text,
    [System.Text.UTF8Encoding]::new($false)
  )
} finally {
  if ($doc) { $doc.Close($false) }
  if ($word) { $word.Quit() }
  if ($doc) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
  if ($word) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
}
```

- [ ] **Step 2: 写解析器夹具和失败测试**

`test/fixtures/question-bank-sample.txt`：

```text
导论 示例
一、判断题
1.马克思主义深刻改变了中国。（ ）
二、单项选择题
1.最新理论成果是（ ）。
A.毛泽东思想 B.科学发展观 C.习近平新时代中国特色社会主义思想 D.邓小平理论
三、多项选择题
1.理论成果包括（ ）。
A.毛泽东思想 B.邓小平理论 C.科学发展观 D.习近平新时代中国特色社会主义思想
参考答案
一、判断题
1.√
二、单项选择题
1.C
三、多项选择题
1.ABCD
```

`test/bank-builder.test.mjs`：

```js
import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { parseQuestionBank } from "../tools/build-question-bank.mjs";

test("解析章节、三种题型和答案", async () => {
  const text = await readFile("test/fixtures/question-bank-sample.txt", "utf8");
  const { questions, report } = parseQuestionBank(text);
  assert.equal(report.sections, 1);
  assert.equal(report.errors.length, 0);
  assert.deepEqual(questions.map(question => question.type), [
    "judgment", "single", "multiple"
  ]);
  assert.deepEqual(questions.map(question => question.answer), [
    ["T"], ["C"], ["A", "B", "C", "D"]
  ]);
});
```

- [ ] **Step 3: 运行并确认失败**

Run: `node --test test/bank-builder.test.mjs`

Expected: FAIL，错误包含 `parseQuestionBank` 不存在。

- [ ] **Step 4: 实现状态机解析器**

`tools/build-question-bank.mjs` 必须导出 `parseQuestionBank(text)`，并在直接运行时接收三个路径参数。实现规则如下：

```js
import { readFile, writeFile } from "node:fs/promises";
import { pathToFileURL } from "node:url";
import { assertQuestion, normalizeSelection } from "../src/domain/questions.js";

const sectionPattern = /^(导论|第[一二三四五六七八九十]+章|结束语)\s*(.*)$/;
const typeHeadings = new Map([
  ["一、判断题", "judgment"],
  ["二、单项选择题", "single"],
  ["三、多项选择题", "multiple"]
]);

function clean(value) {
  return value.replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
}

function splitOptions(body) {
  const matches = [...body.matchAll(/([A-D])[.、]\s*/g)];
  if (matches.length === 0) return { stem: clean(body), options: [] };
  const stem = clean(body.slice(0, matches[0].index));
  const options = matches.map((match, index) => {
    const start = match.index + match[0].length;
    const end = matches[index + 1]?.index ?? body.length;
    return { key: match[1], text: clean(body.slice(start, end)) };
  });
  return { stem, options };
}

function parseNumberedBlocks(lines) {
  const blocks = [];
  let current = null;
  for (const line of lines) {
    const match = line.match(/^(\d+)[.、]\s*(.*)$/);
    if (match) {
      if (current) blocks.push(current);
      current = { number: Number(match[1]), body: match[2] };
    } else if (current) {
      current.body += ` ${line}`;
    }
  }
  if (current) blocks.push(current);
  return blocks;
}

function parseAnswers(lines, type) {
  let joined = lines.join(" ");
  if (type === "judgment" && /^[√×对错]/.test(joined)) {
    joined = `1、${joined}`;
  }
  const pattern = type === "judgment"
    ? /(\d+)[.、]\s*([√×对错])/g
    : /(\d+)[.、]\s*([A-D]{1,4})/gi;
  return new Map([...joined.matchAll(pattern)].map(match => [
    Number(match[1]),
    type === "judgment"
      ? [/[√对]/.test(match[2]) ? "T" : "F"]
      : normalizeSelection([...match[2]])
  ]));
}

function sectionId(index) {
  return `chapter-${String(index).padStart(2, "0")}`;
}

export function parseQuestionBank(text) {
  const lines = text.split(/\r?\n|\r/).map(clean).filter(Boolean);
  const sections = [];
  let current = null;
  for (const line of lines) {
    const section = line.match(sectionPattern);
    if (section) {
      current = { title: clean(`${section[1]} ${section[2]}`), lines: [] };
      sections.push(current);
    } else if (current) {
      current.lines.push(line);
    }
  }

  const questions = [];
  const errors = [];
  sections.forEach((section, sectionIndex) => {
    const answerIndex = section.lines.indexOf("参考答案");
    if (answerIndex < 0) {
      errors.push(`${section.title}: 缺少参考答案`);
      return;
    }
    const questionLines = section.lines.slice(0, answerIndex);
    const answerLines = section.lines.slice(answerIndex + 1);

    for (const [heading, type] of typeHeadings) {
      const qStart = questionLines.indexOf(heading);
      const aStart = answerLines.findIndex(line =>
        line === heading || (type === "judgment" && line === "判断题")
      );
      const nextQuestion = [...typeHeadings.keys()]
        .map(value => questionLines.indexOf(value))
        .filter(index => index > qStart)
        .sort((a, b) => a - b)[0] ?? questionLines.length;
      const nextAnswer = [...typeHeadings.keys()]
        .map(value => answerLines.indexOf(value))
        .filter(index => index > aStart)
        .sort((a, b) => a - b)[0] ?? answerLines.length;
      if (qStart < 0 || aStart < 0) {
        errors.push(`${section.title}: 缺少 ${heading}`);
        continue;
      }
      const qBlocks = parseNumberedBlocks(questionLines.slice(qStart + 1, nextQuestion));
      const answerMap = parseAnswers(answerLines.slice(aStart + 1, nextAnswer), type);

      for (const block of qBlocks) {
        const parsed = type === "judgment"
          ? { stem: clean(block.body.replace(/[（(]\s*[）)]$/, "")), options: [] }
          : splitOptions(block.body);
        const question = {
          id: `${sectionId(sectionIndex)}-${type}-${String(block.number).padStart(3, "0")}`,
          chapterId: sectionId(sectionIndex),
          chapterTitle: section.title,
          type,
          sourceNumber: block.number,
          stem: parsed.stem,
          options: parsed.options,
          answer: answerMap.get(block.number) || [],
          bankVersion: 1
        };
        try {
          assertQuestion(question);
          questions.push(question);
        } catch (error) {
          errors.push(`${question.id}: ${error.message}`);
        }
      }
    }
  });
  const duplicateIds = questions
    .map(question => question.id)
    .filter((id, index, ids) => ids.indexOf(id) !== index);
  errors.push(...duplicateIds.map(id => `${id}: 重复 ID`));
  return {
    questions,
    report: {
      bankVersion: 1,
      sections: sections.length,
      questions: questions.length,
      byType: Object.fromEntries(
        ["judgment", "single", "multiple"].map(type => [
          type, questions.filter(question => question.type === type).length
        ])
      ),
      errors
    }
  };
}

async function main() {
  const [, , input, output, reportOutput] = process.argv;
  if (!input || !output || !reportOutput) {
    throw new Error("用法: node tools/build-question-bank.mjs <input> <output> <report>");
  }
  const result = parseQuestionBank(await readFile(input, "utf8"));
  await writeFile(output, `${JSON.stringify({
    bankVersion: 1,
    questions: result.questions
  }, null, 2)}\n`);
  await writeFile(reportOutput, `${JSON.stringify(result.report, null, 2)}\n`);
  if (result.report.sections !== 10 || result.report.errors.length > 0) {
    process.exitCode = 1;
  }
}

if (import.meta.url === pathToFileURL(process.argv[1] || "").href) {
  await main();
}
```

- [ ] **Step 5: 运行解析器测试**

Run: `node --test test/bank-builder.test.mjs`

Expected: 1 test PASS。

- [ ] **Step 6: 提取实际文档并生成第一版报告**

Run:

```powershell
powershell -ExecutionPolicy Bypass -File tools/extract-doc.ps1 `
  -InputDoc "C:\Users\LGL\Downloads\24毛概练习题整合版.doc" `
  -OutputText "source\question-bank.txt"
npm run bank:build
```

Expected: `data/question-bank-report.json` 中 `sections` 为 `10`。若 `errors` 非空，命令以非零状态退出。

- [ ] **Step 7: 解决报告中的每个异常**

逐条对照源文档修正解析器的格式规则；不得猜测知识性答案。每次修改后运行：

Run: `node --test test/bank-builder.test.mjs && npm run bank:build`

Expected: 测试 PASS，`sections: 10`，`errors: []`，三种题型数量均大于 0。然后人工抽查每个部分的第一题、最后一题和全部多选答案。

- [ ] **Step 8: 提交生成工具和已校验题库**

```bash
git add tools/extract-doc.ps1 tools/build-question-bank.mjs test/fixtures/question-bank-sample.txt test/bank-builder.test.mjs data/questions.v1.json data/question-bank-report.json
git commit -m "feat: import and validate question bank"
```

### Task 4: 练习会话选择器

**Files:**
- Create: `src/domain/session.js`
- Create: `test/session.test.mjs`

- [ ] **Step 1: 写失败测试**

`test/session.test.mjs`：

```js
import test from "node:test";
import assert from "node:assert/strict";
import { createSession } from "../src/domain/session.js";

const questions = Array.from({ length: 6 }, (_, index) => ({
  id: `q${index + 1}`,
  chapterId: index < 3 ? "chapter-00" : "chapter-01",
  type: index % 2 ? "single" : "judgment"
}));

test("章节练习按章节和题型筛选", () => {
  const result = createSession(questions, {
    mode: "chapter",
    chapterId: "chapter-00",
    types: ["single"]
  });
  assert.deepEqual(result.map(item => item.id), ["q2"]);
});

test("随机练习不重复并限制题量", () => {
  const result = createSession(questions, {
    mode: "random",
    count: 3,
    random: () => 0.25
  });
  assert.equal(result.length, 3);
  assert.equal(new Set(result.map(item => item.id)).size, 3);
});

test("请求题量超过可用题数时返回全部", () => {
  assert.equal(createSession(questions, {
    mode: "random",
    count: 50,
    random: () => 0.5
  }).length, 6);
});

test("错题练习只保留传入 ID", () => {
  assert.deepEqual(createSession(questions, {
    mode: "wrong",
    wrongIds: ["q2", "q5"]
  }).map(item => item.id), ["q2", "q5"]);
});
```

- [ ] **Step 2: 运行并确认失败**

Run: `node --test test/session.test.mjs`

Expected: FAIL，错误包含 `ERR_MODULE_NOT_FOUND`。

- [ ] **Step 3: 实现会话选择器**

`src/domain/session.js`：

```js
function shuffled(items, random) {
  const copy = [...items];
  for (let index = copy.length - 1; index > 0; index -= 1) {
    const swap = Math.floor(random() * (index + 1));
    [copy[index], copy[swap]] = [copy[swap], copy[index]];
  }
  return copy;
}

export function createSession(questions, options) {
  const { mode } = options;
  if (mode === "chapter") {
    const types = new Set(options.types || []);
    return questions.filter(question =>
      question.chapterId === options.chapterId &&
      (types.size === 0 || types.has(question.type))
    );
  }
  if (mode === "wrong") {
    const wrongIds = new Set(options.wrongIds || []);
    return questions.filter(question => wrongIds.has(question.id));
  }
  if (mode === "random") {
    const count = Math.max(0, Number(options.count) || 0);
    return shuffled(questions, options.random || Math.random)
      .slice(0, Math.min(count, questions.length));
  }
  throw new Error(`未知练习模式: ${mode}`);
}
```

- [ ] **Step 4: 运行测试并提交**

Run: `node --test test/session.test.mjs`

Expected: 4 tests PASS。

```bash
git add src/domain/session.js test/session.test.mjs
git commit -m "feat: add chapter random and wrong sessions"
```

### Task 5: 学习记录、错题规则和统计

**Files:**
- Create: `src/domain/progress.js`
- Create: `test/progress.test.mjs`

- [ ] **Step 1: 写失败测试**

`test/progress.test.mjs`：

```js
import test from "node:test";
import assert from "node:assert/strict";
import { applyAttempt, buildStats } from "../src/domain/progress.js";

test("答错进入错题本并累计错误次数", () => {
  const record = applyAttempt(undefined, {
    questionId: "q1", selected: ["A"], correct: false, at: "2026-07-02T10:00:00.000Z"
  });
  assert.equal(record.attempts, 1);
  assert.equal(record.errors, 1);
  assert.equal(record.isWrong, true);
  assert.equal(record.mastered, false);
});

test("在错题模式答对后标记已掌握", () => {
  const wrong = applyAttempt(undefined, {
    questionId: "q1", selected: ["A"], correct: false, at: "2026-07-02T10:00:00.000Z"
  });
  const mastered = applyAttempt(wrong, {
    questionId: "q1", selected: ["C"], correct: true,
    mode: "wrong", at: "2026-07-02T10:05:00.000Z"
  });
  assert.equal(mastered.isWrong, false);
  assert.equal(mastered.mastered, true);
  assert.equal(mastered.errors, 1);
});

test("普通练习再次答错会重新进入错题本", () => {
  const mastered = { questionId: "q1", attempts: 2, correct: 1, errors: 1, isWrong: false, mastered: true };
  const result = applyAttempt(mastered, {
    questionId: "q1", selected: ["B"], correct: false, mode: "chapter",
    at: "2026-07-02T11:00:00.000Z"
  });
  assert.equal(result.isWrong, true);
  assert.equal(result.mastered, false);
});

test("统计已练数量、正确率和错题数", () => {
  const stats = buildStats([
    { attempts: 2, correct: 1, isWrong: true },
    { attempts: 1, correct: 1, isWrong: false }
  ], 10);
  assert.deepEqual(stats, {
    totalQuestions: 10, practiced: 2, attempts: 3,
    correct: 2, accuracy: 67, wrong: 1
  });
});
```

- [ ] **Step 2: 运行并确认失败**

Run: `node --test test/progress.test.mjs`

Expected: FAIL，错误包含 `ERR_MODULE_NOT_FOUND`。

- [ ] **Step 3: 实现学习记录纯函数**

`src/domain/progress.js`：

```js
export function applyAttempt(previous, attempt) {
  const base = previous || {
    questionId: attempt.questionId,
    attempts: 0,
    correct: 0,
    errors: 0,
    isWrong: false,
    mastered: false
  };
  const correct = Boolean(attempt.correct);
  const mastered = correct && attempt.mode === "wrong";
  return {
    ...base,
    attempts: base.attempts + 1,
    correct: base.correct + (correct ? 1 : 0),
    errors: base.errors + (correct ? 0 : 1),
    isWrong: correct ? (mastered ? false : base.isWrong) : true,
    mastered: correct ? (mastered || base.mastered) : false,
    lastSelected: [...attempt.selected],
    lastCorrect: correct,
    lastAttemptAt: attempt.at
  };
}

export function buildStats(records, totalQuestions) {
  const practiced = records.filter(record => record.attempts > 0).length;
  const attempts = records.reduce((sum, record) => sum + record.attempts, 0);
  const correct = records.reduce((sum, record) => sum + record.correct, 0);
  return {
    totalQuestions,
    practiced,
    attempts,
    correct,
    accuracy: attempts === 0 ? 0 : Math.round((correct / attempts) * 100),
    wrong: records.filter(record => record.isWrong).length
  };
}
```

- [ ] **Step 4: 运行测试并提交**

Run: `node --test test/progress.test.mjs`

Expected: 4 tests PASS。

```bash
git add src/domain/progress.js test/progress.test.mjs
git commit -m "feat: track progress and wrong answers"
```

### Task 6: IndexedDB 持久化和事务式备份

**Files:**
- Create: `src/data/progress-store.js`
- Create: `src/data/backup.js`
- Create: `test/backup.test.mjs`

- [ ] **Step 1: 写备份失败测试**

`test/backup.test.mjs`：

```js
import test from "node:test";
import assert from "node:assert/strict";
import { createBackup, parseBackup } from "../src/data/backup.js";

test("备份包含版本、题库版本和记录", () => {
  const text = createBackup([{ questionId: "q1", attempts: 1 }], {
    exportedAt: "2026-07-02T12:00:00.000Z",
    bankVersion: 1
  });
  assert.deepEqual(parseBackup(text, 1).records, [{ questionId: "q1", attempts: 1 }]);
});

test("拒绝错误格式和较新题库版本", () => {
  assert.throws(() => parseBackup("{}", 1), /备份格式/);
  const newer = createBackup([], {
    exportedAt: "2026-07-02T12:00:00.000Z",
    bankVersion: 2
  });
  assert.throws(() => parseBackup(newer, 1), /题库版本/);
});
```

- [ ] **Step 2: 实现备份编码和校验**

`src/data/backup.js`：

```js
import { BACKUP_FORMAT, BACKUP_VERSION } from "../config.js";

export function createBackup(records, { exportedAt, bankVersion }) {
  return `${JSON.stringify({
    format: BACKUP_FORMAT,
    version: BACKUP_VERSION,
    bankVersion,
    exportedAt,
    records
  }, null, 2)}\n`;
}

export function parseBackup(text, currentBankVersion) {
  let value;
  try {
    value = JSON.parse(text);
  } catch {
    throw new Error("备份文件不是有效 JSON");
  }
  if (value.format !== BACKUP_FORMAT || value.version !== BACKUP_VERSION) {
    throw new Error("备份格式或版本不受支持");
  }
  if (!Number.isInteger(value.bankVersion) || value.bankVersion > currentBankVersion) {
    throw new Error("备份题库版本高于当前应用");
  }
  if (!Array.isArray(value.records) ||
      value.records.some(record => !record || typeof record.questionId !== "string")) {
    throw new Error("备份记录结构无效");
  }
  return value;
}
```

- [ ] **Step 3: 运行备份测试**

Run: `node --test test/backup.test.mjs`

Expected: 2 tests PASS。

- [ ] **Step 4: 实现 IndexedDB 存储**

`src/data/progress-store.js`：

```js
import { DB_NAME, DB_VERSION } from "../config.js";

function requestResult(request) {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

export async function openProgressStore(indexedDBFactory = indexedDB) {
  const request = indexedDBFactory.open(DB_NAME, DB_VERSION);
  request.onupgradeneeded = () => {
    const database = request.result;
    if (!database.objectStoreNames.contains("progress")) {
      database.createObjectStore("progress", { keyPath: "questionId" });
    }
  };
  const database = await requestResult(request);

  function store(mode) {
    return database.transaction("progress", mode).objectStore("progress");
  }

  return {
    async all() {
      return requestResult(store("readonly").getAll());
    },
    async get(questionId) {
      return requestResult(store("readonly").get(questionId));
    },
    async put(record) {
      await requestResult(store("readwrite").put(record));
    },
    async replaceAll(records) {
      await new Promise((resolve, reject) => {
        const transaction = database.transaction("progress", "readwrite");
        const objectStore = transaction.objectStore("progress");
        objectStore.clear();
        records.forEach(record => objectStore.put(record));
        transaction.oncomplete = resolve;
        transaction.onerror = () => reject(transaction.error);
        transaction.onabort = () => reject(transaction.error);
      });
    },
    async clear() {
      await requestResult(store("readwrite").clear());
    },
    close() {
      database.close();
    }
  };
}
```

- [ ] **Step 5: 提交**

```bash
git add src/data/progress-store.js src/data/backup.js test/backup.test.mjs
git commit -m "feat: persist progress and support backups"
```

### Task 7: 题库仓库、页面渲染与完整答题流程

**Files:**
- Create: `src/data/question-repository.js`
- Create: `src/ui/screens.js`
- Create: `src/ui/controller.js`
- Create: `src/app.js`
- Create: `styles/tokens.css`
- Create: `styles/base.css`
- Create: `styles/components.css`

- [ ] **Step 1: 实现题库仓库**

`src/data/question-repository.js`：

```js
import { assertQuestion } from "../domain/questions.js";

export async function loadQuestionBank(url, expectedVersion, fetcher = fetch) {
  const response = await fetcher(url);
  if (!response.ok) throw new Error(`题库加载失败：HTTP ${response.status}`);
  const bank = await response.json();
  if (bank.bankVersion !== expectedVersion || !Array.isArray(bank.questions)) {
    throw new Error("题库版本或结构无效");
  }
  bank.questions.forEach(assertQuestion);
  return bank.questions;
}
```

- [ ] **Step 2: 创建纯渲染函数**

`src/ui/screens.js` 导出以下函数，所有动态文本先经过 `escapeHtml`：

```js
export function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, character => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  })[character]);
}

export function renderHome(stats) {
  return `<main class="screen">
    <header class="hero"><p class="eyebrow">毛概刷题</p><h1>继续复习</h1>
      <p>距离期末更从容一点</p></header>
    <section class="stats" aria-label="学习统计">
      <div class="stat-card"><strong>${stats.practiced}</strong><span>已练</span></div>
      <div class="stat-card"><strong>${stats.accuracy}%</strong><span>正确率</span></div>
      <div class="stat-card"><strong>${stats.wrong}</strong><span>错题</span></div>
    </section>
    <nav class="home-actions">
      <button data-action="open-chapters">章节练习</button>
      <button data-action="open-random">随机练习</button>
      <button data-action="open-wrong">复习错题</button>
    </nav>
  </main>${renderBottomNav("home")}`;
}

export function renderBottomNav(active) {
  return `<nav class="bottom-nav" aria-label="主导航">
    ${[["home","首页"],["practice","练习"],["wrong","错题"],["settings","设置"]]
      .map(([route, label]) => `<button data-route="${route}" aria-current="${active === route ? "page" : "false"}">${label}</button>`)
      .join("")}
  </nav>`;
}

export function renderQuiz({ question, index, total, selected, locked, correct }) {
  const keys = question.type === "judgment" ? [
    { key: "T", text: "正确" }, { key: "F", text: "错误" }
  ] : question.options;
  return `<main class="screen quiz-screen">
    <header class="quiz-header"><button data-action="quit" aria-label="退出练习">‹</button>
      <div><strong>${escapeHtml(question.chapterTitle)}</strong><small>第 ${index + 1} / ${total} 题</small></div>
      <button data-action="bookmark" aria-label="收藏题目">☆</button></header>
    <progress max="${total}" value="${index + 1}"></progress>
    <article class="question-card">
      <span class="type-tag">${{judgment:"判断题",single:"单项选择题",multiple:"多项选择题"}[question.type]}</span>
      <h1>${escapeHtml(question.stem)}</h1>
      <fieldset ${locked ? "disabled" : ""}>
        <legend class="sr-only">请选择答案</legend>
        ${keys.map(option => `<label class="answer-option">
          <input type="${question.type === "multiple" ? "checkbox" : "radio"}"
            name="answer" value="${option.key}" ${selected.includes(option.key) ? "checked" : ""}>
          <span class="option-key">${option.key}</span><span>${escapeHtml(option.text)}</span>
        </label>`).join("")}
      </fieldset>
      ${locked ? `<div class="feedback ${correct ? "correct" : "wrong"}">
        <strong>${correct ? "回答正确" : "回答错误"}</strong>
        <span>标准答案：${question.answer.join("")}</span></div>` : ""}
    </article>
    <button class="primary fixed-action" data-action="${locked ? "next" : "submit"}">
      ${locked ? "下一题" : "提交答案"}</button>
  </main>`;
}
```

同一文件还需导出以下固定接口：

- `renderChapterPicker(chapters)`
- `renderRandomPicker()`
- `renderResults(summary)`
- `renderWrongList(questions, records)`
- `renderSettings({ appVersion, bankVersion, persistent })`
- `renderError(message)`

这些函数只生成已在设计规格中确认的控件，不增加模拟考试、解析、登录或云同步入口。

- [ ] **Step 3: 实现控制器和应用状态流**

`src/ui/controller.js`：

```js
export function readSelectedAnswers(root = document) {
  return [...root.querySelectorAll('input[name="answer"]:checked')]
    .map(input => input.value);
}

export function bindAppActions(root, handlers) {
  root.addEventListener("click", event => {
    const target = event.target.closest("[data-action],[data-route]");
    if (!target) return;
    if (target.dataset.action) handlers.action(target.dataset.action, target);
    if (target.dataset.route) handlers.route(target.dataset.route);
  });
}
```

`src/app.js` 的状态机按以下固定顺序实现：

1. 并行载入题库和打开 IndexedDB。
2. IndexedDB 失败时切换为内存存储，并显示“不保存记录”横幅。
3. 根据 `location.hash` 渲染首页、选择页、答题页、结果页、错题页和设置页。
4. 提交时调用 `gradeQuestion`，再调用 `applyAttempt` 和 `store.put`。
5. 结果页使用本次会话内的结果，不从 DOM 反推数据。
6. 导入备份时先调用 `parseBackup`，确认后调用 `replaceAll`。
7. 清空记录时使用原生确认对话框二次确认。

核心状态固定为：

```js
const state = {
  questions: [],
  records: [],
  store: null,
  persistent: true,
  session: null,
  index: 0,
  selected: [],
  locked: false,
  lastCorrect: null,
  sessionResults: []
};
```

应用协调器只调用已定义模块，不重复判分、随机或统计逻辑。

- [ ] **Step 4: 实现清爽学习蓝样式**

`styles/tokens.css`：

```css
:root {
  --bg: #f6f8ff;
  --surface: #ffffff;
  --text: #172554;
  --muted: #64748b;
  --primary: #2563eb;
  --primary-soft: #e0e7ff;
  --success: #15803d;
  --success-soft: #f0fdf4;
  --danger: #dc2626;
  --danger-soft: #fef2f2;
  --border: #dbeafe;
  --radius-sm: 12px;
  --radius-md: 16px;
  --radius-lg: 22px;
  --shadow: 0 8px 24px rgb(37 99 235 / 8%);
  --space-1: 4px;
  --space-2: 8px;
  --space-3: 12px;
  --space-4: 16px;
  --space-5: 20px;
  --space-6: 24px;
}
```

`styles/base.css` 和 `styles/components.css` 的验收规则为：

- `box-sizing: border-box`
- 系统字体栈
- `body` 最小宽度 320px、背景 `--bg`
- `.screen` 最大宽度 640px 并居中
- 所有按钮最小高度 44px
- `.sr-only` 可供读屏但视觉隐藏
- 卡片、统计、选项、反馈、进度条和固定底部按钮
- `padding-bottom: calc(88px + env(safe-area-inset-bottom))`
- `:focus-visible` 使用 3px 蓝色轮廓
- 360px 宽度无横向滚动

- [ ] **Step 5: 运行全部自动测试**

Run: `npm test`

Expected: 题目、解析器、会话、进度和备份测试全部 PASS。

- [ ] **Step 6: 浏览器验证完整流程**

Run: `npm run serve`

在 360×800 手机视口依次验证：

1. 首页统计为零。
2. 章节练习可以选择章节和题型。
3. 判断、单选、多选提交后立即判题。
4. 多选少选判错。
5. 错题自动进入错题本。
6. 错题本答对后移出当前列表。
7. 刷新页面后记录仍存在。
8. 导出备份、清空、导入后统计恢复一致。

Expected: 控制台无未处理异常，页面无横向滚动，主要按钮单手可触达。

- [ ] **Step 7: 提交完整业务流程**

```bash
git add src styles
git commit -m "feat: build mobile quiz experience"
```

### Task 8: PWA 安装、图标与离线更新

**Files:**
- Create: `manifest.webmanifest`
- Create: `service-worker.js`
- Create: `src/ui/install.js`
- Create: `assets/icons/icon-192.png`
- Create: `assets/icons/icon-512.png`
- Modify: `src/app.js`

- [ ] **Step 1: 创建安装清单**

`manifest.webmanifest`：

```json
{
  "name": "毛概期末刷题",
  "short_name": "毛概刷题",
  "description": "可离线使用的毛概期末判断、单选和多选练习",
  "lang": "zh-CN",
  "start_url": "./",
  "scope": "./",
  "display": "standalone",
  "background_color": "#f6f8ff",
  "theme_color": "#2563eb",
  "icons": [
    {
      "src": "./assets/icons/icon-192.png",
      "sizes": "192x192",
      "type": "image/png",
      "purpose": "any maskable"
    },
    {
      "src": "./assets/icons/icon-512.png",
      "sizes": "512x512",
      "type": "image/png",
      "purpose": "any maskable"
    }
  ]
}
```

- [ ] **Step 2: 生成两种尺寸的应用图标**

图标使用已确认的学习蓝：蓝色圆角方形底，中央为白色书本与勾号，四周保留 20% maskable 安全区。输出必须为真实的 192×192 和 512×512 PNG，并用图像检查工具确认尺寸。

- [ ] **Step 3: 实现版本化 Service Worker**

`service-worker.js`：

```js
const CACHE = "maogai-quiz-v1";
const ASSETS = [
  "./",
  "./index.html",
  "./manifest.webmanifest",
  "./styles/tokens.css",
  "./styles/base.css",
  "./styles/components.css",
  "./src/app.js",
  "./src/config.js",
  "./src/domain/questions.js",
  "./src/domain/session.js",
  "./src/domain/progress.js",
  "./src/data/question-repository.js",
  "./src/data/progress-store.js",
  "./src/data/backup.js",
  "./src/ui/screens.js",
  "./src/ui/controller.js",
  "./src/ui/install.js",
  "./data/questions.v1.json",
  "./assets/icons/icon-192.png",
  "./assets/icons/icon-512.png"
];

self.addEventListener("install", event => {
  event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(ASSETS)));
});

self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(key => key !== CACHE).map(key => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", event => {
  if (event.request.method !== "GET") return;
  event.respondWith(
    caches.match(event.request).then(cached =>
      cached || fetch(event.request).then(response => {
        if (!response.ok) return response;
        const copy = response.clone();
        caches.open(CACHE).then(cache => cache.put(event.request, copy));
        return response;
      })
    )
  );
});
```

- [ ] **Step 4: 实现安装提示**

`src/ui/install.js`：

```js
let deferredPrompt = null;

export function watchInstallPrompt(onAvailable) {
  window.addEventListener("beforeinstallprompt", event => {
    event.preventDefault();
    deferredPrompt = event;
    onAvailable(true);
  });
}

export async function requestInstall() {
  if (!deferredPrompt) return false;
  deferredPrompt.prompt();
  const result = await deferredPrompt.userChoice;
  deferredPrompt = null;
  return result.outcome === "accepted";
}

export function installGuidance() {
  const ua = navigator.userAgent;
  if (/MicroMessenger/i.test(ua)) return "请使用右上角菜单，在系统浏览器中打开后安装。";
  if (/iPhone|iPad/i.test(ua)) return "请点按 Safari 分享按钮，再选择“添加到主屏幕”。";
  return "请使用浏览器菜单中的“安装应用”或“添加到主屏幕”。";
}
```

在 `src/app.js` 启动结束后注册：

```js
if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => navigator.serviceWorker.register("./service-worker.js"));
}
```

- [ ] **Step 5: 验证离线和安装**

Run: `npm run serve`

Expected:

1. Manifest 可加载且图标无错误。
2. Service Worker 安装后缓存列表完整。
3. 首次在线加载后切换浏览器离线模式，刷新仍能进入首页、读取题库并答题。
4. IndexedDB 记录在 Service Worker 更新后仍保留。
5. 安卓 Chrome 可触发安装；iPhone 和微信显示对应文字说明。

- [ ] **Step 6: 提交 PWA 能力**

```bash
git add manifest.webmanifest service-worker.js src/ui/install.js src/app.js assets/icons
git commit -m "feat: add installable offline PWA"
```

### Task 9: 最终回归、题库抽查和静态部署

**Files:**
- Modify: `README.md`
- Modify: `data/question-bank-report.json`（仅当最终重建结果变化）

- [ ] **Step 1: 运行完整自动测试和题库重建**

Run:

```bash
npm test
npm run bank:build
git diff --exit-code data/questions.v1.json data/question-bank-report.json
```

Expected: 所有测试 PASS；题库报告 `sections: 10`、`errors: []`；重新生成后无差异。

- [ ] **Step 2: 完成题库人工抽查**

对 10 个部分逐一检查：

- 每部分判断题、单选题、多选题各抽查第一题和最后一题
- 所有报告曾标记过异常的题目
- 所有多选答案是否只包含 A-D 且与源文档一致
- 首页显示的总题数是否等于报告中的 `questions`

Expected: 抽查项全部与源文档一致。任何差异回到 Task 3 修复解析规则并重新生成。

- [ ] **Step 3: 完成手机回归矩阵**

在 360×800、390×844、430×932 三种视口验证首页、答题、反馈、结果、错题和设置；在实际安卓 Chrome 验证安装与离线；在 iPhone Safari 验证“添加到主屏幕”说明。

Expected: 无遮挡、横向滚动、不可触达按钮或仅用颜色表达的状态。

- [ ] **Step 4: 写 README**

`README.md` 使用以下完整内容：

````markdown
# 毛概期末刷题

可安装、可离线、无需账号的手机端毛概判断题、单选题和多选题练习应用。

## 本地运行

```powershell
npm run serve
```

打开 http://127.0.0.1:4173。

## 测试

```powershell
npm test
```

## 重新构建题库

将原始 Word 文档放在本机后运行：

```powershell
powershell -ExecutionPolicy Bypass -File tools/extract-doc.ps1 `
  -InputDoc "原始文档路径.doc" `
  -OutputText "source\question-bank.txt"
npm run bank:build
```

只有 `data/question-bank-report.json` 中 `sections` 为 10 且 `errors` 为空时，题库构建才算通过。

## 部署

仓库根目录就是静态站点根目录。部署到 GitHub Pages 后，首次在线打开完成缓存，即可离线使用。

## 数据说明

学习记录保存在当前浏览器的 IndexedDB 中。清理浏览器数据会删除记录，请定期在设置页导出备份。
````

- [ ] **Step 5: 检查 GitHub Pages 路径兼容性**

使用子路径形式访问站点，确认所有链接均为 `./` 相对路径，不出现以 `/` 开头的静态资源路径。

Expected: 部署在 `https://<user>.github.io/<repo>/` 时，HTML、脚本、样式、题库、图标和 Service Worker 全部返回 200。

- [ ] **Step 6: 最终提交**

```bash
git add README.md data/question-bank-report.json
git commit -m "docs: add usage deployment and QA guide"
git status --short
```

Expected: 工作区干净。

## 实施完成检查表

- [ ] `npm test` 全部通过
- [ ] 题库 10 个部分完整，异常数为 0
- [ ] 三种题型判分与源答案一致
- [ ] 章节、随机、错题流程完整
- [ ] 刷新和重启后记录保留
- [ ] 备份导出、清空、导入恢复一致
- [ ] 首次加载后完全离线可用
- [ ] 安卓可安装，iPhone 有明确安装指引
- [ ] 360px 宽度无布局缺陷
- [ ] Git 工作区干净且每个任务独立提交

import { readFile, mkdir, writeFile } from "node:fs/promises";
import { dirname } from "node:path";
import { pathToFileURL } from "node:url";
import { assertQuestion, normalizeSelection } from "../src/domain/questions.js";

const SECTION = /^(导论|第[一二三四五六七八九十]+章|结束语)\s*(.*)$/;
const HEADINGS = [
  ["一、判断题", "judgment"],
  ["二、单项选择题", "single"],
  ["三、多项选择题", "multiple"]
];

function clean(value) {
  return value.replace(/\u00a0/g, " ").replace(/[\t ]+/g, " ").trim();
}

function parseNumberedBlocks(lines) {
  const blocks = [];
  let current;
  for (const line of lines) {
    const match = line.match(/^(\d+)(?:[.、．]\s*|\s*)(.*)$/);
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

function splitOptions(body) {
  const matches = [...body.matchAll(/([A-D])(?:[.、．]\s*|(?=[\u3400-\u9fff]))/g)];
  if (!matches.length) return { stem: clean(body), options: [] };
  return {
    stem: clean(body.slice(0, matches[0].index)),
    options: matches.map((match, index) => ({
      key: match[1],
      text: clean(body.slice(match.index + match[0].length, matches[index + 1]?.index ?? body.length))
    }))
  };
}

function parseAnswers(lines, type) {
  let joined = lines.join(" ");
  if (type === "judgment" && /^[√×对错]/.test(joined)) joined = `1、${joined}`;
  const pattern = type === "judgment"
    ? /(\d+)[.、．]\s*([√×对错])/g
    : /(\d+)[.、．]\s*([A-D]{1,4})/gi;
  return new Map([...joined.matchAll(pattern)].map(match => [
    Number(match[1]),
    type === "judgment"
      ? [/[√对]/.test(match[2]) ? "T" : "F"]
      : normalizeSelection([...match[2]])
  ]));
}

function nextHeadingIndex(lines, start) {
  return HEADINGS.map(([heading]) => lines.indexOf(heading))
    .filter(index => index > start).sort((a, b) => a - b)[0] ?? lines.length;
}

export function parseQuestionBank(text, { expectedSections = 10 } = {}) {
  const lines = text.split(/\r\n|\r|\n/).map(clean).filter(Boolean);
  const sections = [];
  let current;
  for (const line of lines) {
    const match = line.match(SECTION);
    if (match) {
      current = { title: clean(`${match[1]} ${match[2]}`), lines: [] };
      sections.push(current);
    } else if (current) {
      current.lines.push(line);
    }
  }

  const questions = [];
  const errors = [];
  if (sections.length !== expectedSections) errors.push(`章节数应为 ${expectedSections}，实际为 ${sections.length}`);
  sections.forEach((section, sectionIndex) => {
    const answerIndex = section.lines.indexOf("参考答案");
    if (answerIndex < 0) {
      errors.push(`${section.title}: 缺少参考答案`);
      return;
    }
    const questionLines = section.lines.slice(0, answerIndex);
    const answerLines = section.lines.slice(answerIndex + 1);
    for (const [heading, type] of HEADINGS) {
      const qStart = questionLines.indexOf(heading);
      const answerNames = type === "judgment" ? [heading, "判断题"] : [heading];
      const aStart = answerLines.findIndex(line => answerNames.includes(line));
      if (qStart < 0 || aStart < 0) {
        errors.push(`${section.title}: 缺少 ${heading}`);
        continue;
      }
      const qBlocks = parseNumberedBlocks(questionLines.slice(qStart + 1, nextHeadingIndex(questionLines, qStart)));
      const answers = parseAnswers(answerLines.slice(aStart + 1, nextHeadingIndex(answerLines, aStart)), type);
      for (const block of qBlocks) {
        const parsed = type === "judgment"
          ? { stem: clean(block.body.replace(/[（(]\s*[）)]\s*$/, "")), options: [] }
          : splitOptions(block.body);
        const question = {
          id: `chapter-${String(sectionIndex).padStart(2, "0")}-${type}-${String(block.number).padStart(3, "0")}`,
          chapterId: `chapter-${String(sectionIndex).padStart(2, "0")}`,
          chapterTitle: section.title,
          type,
          sourceNumber: block.number,
          stem: parsed.stem,
          options: parsed.options,
          answer: answers.get(block.number) || [],
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

  const ids = questions.map(question => question.id);
  ids.filter((id, index) => ids.indexOf(id) !== index)
    .forEach(id => errors.push(`${id}: 重复 ID`));
  return {
    questions,
    report: {
      bankVersion: 1,
      sections: sections.length,
      questions: questions.length,
      byType: Object.fromEntries(["judgment", "single", "multiple"].map(type => [
        type, questions.filter(question => question.type === type).length
      ])),
      errors
    }
  };
}

async function main() {
  const [, , input, output, reportOutput] = process.argv;
  if (!input || !output || !reportOutput) throw new Error("用法: build-question-bank <input> <output> <report>");
  const result = parseQuestionBank(await readFile(input, "utf8"));
  await mkdir(dirname(output), { recursive: true });
  await writeFile(output, `${JSON.stringify({ bankVersion: 1, questions: result.questions }, null, 2)}\n`);
  await writeFile(reportOutput, `${JSON.stringify(result.report, null, 2)}\n`);
  if (result.report.errors.length) process.exitCode = 1;
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) await main();

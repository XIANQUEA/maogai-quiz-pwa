import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { parseQuestionBank } from "../tools/build-question-bank.mjs";

test("解析章节、三种题型和首题省略编号的答案", async () => {
  const text = await readFile("test/fixtures/question-bank-sample.txt", "utf8");
  const { questions, report } = parseQuestionBank(text, { expectedSections: 1 });
  assert.equal(report.sections, 1);
  assert.equal(report.errors.length, 0);
  assert.deepEqual(questions.map(question => question.type), ["judgment", "single", "multiple"]);
  assert.deepEqual(questions.map(question => question.answer), [["T"], ["C"], ["A", "B", "C", "D"]]);
});

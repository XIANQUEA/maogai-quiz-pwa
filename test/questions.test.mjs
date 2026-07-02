import test from "node:test";
import assert from "node:assert/strict";
import { assertQuestion, gradeQuestion, normalizeSelection } from "../src/domain/questions.js";

const base = {
  id: "chapter-00-single-004", chapterId: "chapter-00", chapterTitle: "导论",
  type: "single", sourceNumber: 4, stem: "最新理论成果是？",
  options: [{ key: "A", text: "毛泽东思想" }, { key: "C", text: "习近平新时代中国特色社会主义思想" }],
  answer: ["C"], bankVersion: 1
};

test("单选题答案一致时正确", () => {
  assert.equal(gradeQuestion(base, ["C"]), true);
  assert.equal(gradeQuestion(base, ["A"]), false);
});

test("多选题忽略顺序但不允许少选", () => {
  const question = { ...base, type: "multiple", answer: ["A", "C"] };
  assert.equal(gradeQuestion(question, ["C", "A"]), true);
  assert.equal(gradeQuestion(question, ["A"]), false);
});

test("判断题只接受 T 或 F", () => {
  const question = { ...base, type: "judgment", options: [], answer: ["T"] };
  assert.equal(gradeQuestion(question, ["T"]), true);
  assert.throws(() => assertQuestion({ ...question, answer: ["A"] }), /判断题答案/);
});

test("选择值会去重、排序并转大写", () => {
  assert.deepEqual(normalizeSelection(["c", "A", "c"]), ["A", "C"]);
});

test("答案引用不存在选项时拒绝题目", () => {
  assert.throws(() => assertQuestion({ ...base, answer: ["D"] }), /不存在选项/);
});

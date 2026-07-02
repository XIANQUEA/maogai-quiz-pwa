import test from "node:test";
import assert from "node:assert/strict";
import { createSession } from "../src/domain/session.js";

const questions = Array.from({ length: 6 }, (_, index) => ({
  id: `q${index + 1}`, chapterId: index < 3 ? "chapter-00" : "chapter-01",
  type: index % 2 ? "single" : "judgment"
}));

test("章节练习按章节和题型筛选", () => {
  assert.deepEqual(createSession(questions, {
    mode: "chapter", chapterId: "chapter-00", types: ["single"]
  }).map(item => item.id), ["q2"]);
});

test("随机练习不重复并限制题量", () => {
  const result = createSession(questions, { mode: "random", count: 3, random: () => 0.25 });
  assert.equal(result.length, 3);
  assert.equal(new Set(result.map(item => item.id)).size, 3);
});

test("请求题量超过题库时返回全部", () => {
  assert.equal(createSession(questions, { mode: "random", count: 50, random: () => 0.5 }).length, 6);
});

test("错题练习只保留传入 ID", () => {
  assert.deepEqual(createSession(questions, {
    mode: "wrong", wrongIds: ["q2", "q5"]
  }).map(item => item.id), ["q2", "q5"]);
});

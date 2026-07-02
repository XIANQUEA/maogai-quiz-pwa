import test from "node:test";
import assert from "node:assert/strict";
import { applyAttempt, buildStats } from "../src/domain/progress.js";

test("答错进入错题本并累计次数", () => {
  const result = applyAttempt(undefined, {
    questionId: "q1", selected: ["A"], isCorrect: false, mode: "chapter", at: "2026-07-02T10:00:00Z"
  });
  assert.deepEqual([result.attempts, result.errorCount, result.isWrong, result.mastered], [1, 1, true, false]);
});

test("错题模式答对后掌握，普通模式再错后重新进入", () => {
  const wrong = applyAttempt(undefined, { questionId: "q1", selected: ["A"], isCorrect: false, mode: "chapter", at: "a" });
  const mastered = applyAttempt(wrong, { questionId: "q1", selected: ["C"], isCorrect: true, mode: "wrong", at: "b" });
  assert.deepEqual([mastered.isWrong, mastered.mastered, mastered.errorCount], [false, true, 1]);
  const again = applyAttempt(mastered, { questionId: "q1", selected: ["B"], isCorrect: false, mode: "random", at: "c" });
  assert.deepEqual([again.isWrong, again.mastered], [true, false]);
});

test("统计已练、正确率和错题数", () => {
  assert.deepEqual(buildStats([
    { attempts: 2, correctCount: 1, isWrong: true },
    { attempts: 1, correctCount: 1, isWrong: false }
  ], 10), { totalQuestions: 10, practiced: 2, attempts: 3, correct: 2, accuracy: 67, wrong: 1 });
});

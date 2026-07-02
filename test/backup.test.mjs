import test from "node:test";
import assert from "node:assert/strict";
import { createBackup, parseBackup } from "../src/data/backup.js";

test("备份可往返恢复记录", () => {
  const text = createBackup([{ questionId: "q1", attempts: 1 }], {
    exportedAt: "2026-07-02T12:00:00Z", bankVersion: 1
  });
  assert.deepEqual(parseBackup(text, 1).records, [{ questionId: "q1", attempts: 1 }]);
});

test("拒绝错误格式和较新题库版本", () => {
  assert.throws(() => parseBackup("{}", 1), /备份格式/);
  const newer = createBackup([], { exportedAt: "now", bankVersion: 2 });
  assert.throws(() => parseBackup(newer, 1), /题库版本/);
});

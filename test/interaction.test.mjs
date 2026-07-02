import test from "node:test";
import assert from "node:assert/strict";
import { shouldSubmitOnSelection } from "../src/ui/controller.js";

test("判断题和单选题选择后立即判题", () => {
  assert.equal(shouldSubmitOnSelection("judgment"), true);
  assert.equal(shouldSubmitOnSelection("single"), true);
});

test("多选题选择后等待用户提交", () => {
  assert.equal(shouldSubmitOnSelection("multiple"), false);
});

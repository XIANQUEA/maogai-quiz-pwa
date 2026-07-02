const TYPES = new Set(["judgment", "single", "multiple"]);

export function normalizeSelection(values = []) {
  return [...new Set(values.map(value => String(value).trim().toUpperCase()))].sort();
}

export function assertQuestion(question) {
  if (!question || typeof question !== "object") throw new Error("题目必须是对象");
  for (const key of ["id", "chapterId", "chapterTitle", "type", "stem"]) {
    if (!String(question[key] ?? "").trim()) throw new Error(`题目缺少 ${key}`);
  }
  if (!TYPES.has(question.type)) throw new Error(`未知题型 ${question.type}`);
  if (!Array.isArray(question.answer) || question.answer.length === 0) throw new Error("题目缺少答案");
  const answer = normalizeSelection(question.answer);
  if (question.type === "judgment") {
    if (answer.length !== 1 || !["T", "F"].includes(answer[0])) throw new Error("判断题答案必须是 T 或 F");
    return true;
  }
  const options = Array.isArray(question.options) ? question.options : [];
  const keys = new Set(options.map(option => option.key));
  if (keys.size !== options.length || options.length < 2) throw new Error("选择题选项无效");
  for (const value of answer) {
    if (!keys.has(value)) throw new Error(`答案 ${value} 不存在选项`);
  }
  if (question.type === "single" && answer.length !== 1) throw new Error("单选题只能有一个答案");
  return true;
}

export function gradeQuestion(question, selected) {
  assertQuestion(question);
  const actual = normalizeSelection(selected);
  const expected = normalizeSelection(question.answer);
  return actual.length === expected.length && actual.every((value, index) => value === expected[index]);
}

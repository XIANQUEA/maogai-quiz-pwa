import { assertQuestion } from "../domain/questions.js";

export async function loadQuestionBank(url, expectedVersion, fetcher = fetch) {
  const response = await fetcher(url);
  if (!response.ok) throw new Error(`题库加载失败：HTTP ${response.status}`);
  const bank = await response.json();
  if (bank.bankVersion !== expectedVersion || !Array.isArray(bank.questions)) throw new Error("题库版本或结构无效");
  bank.questions.forEach(assertQuestion);
  return bank.questions;
}

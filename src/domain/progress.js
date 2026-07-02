export function applyAttempt(previous, attempt) {
  const base = previous || {
    questionId: attempt.questionId, attempts: 0, correctCount: 0,
    errorCount: 0, isWrong: false, mastered: false
  };
  const mastered = attempt.isCorrect && attempt.mode === "wrong";
  return {
    ...base,
    attempts: base.attempts + 1,
    correctCount: base.correctCount + (attempt.isCorrect ? 1 : 0),
    errorCount: base.errorCount + (attempt.isCorrect ? 0 : 1),
    isWrong: attempt.isCorrect ? (mastered ? false : base.isWrong) : true,
    mastered: attempt.isCorrect ? (mastered || base.mastered) : false,
    lastSelected: [...attempt.selected],
    lastCorrect: attempt.isCorrect,
    lastAttemptAt: attempt.at
  };
}

export function buildStats(records, totalQuestions) {
  const practiced = records.filter(record => record.attempts > 0).length;
  const attempts = records.reduce((sum, record) => sum + record.attempts, 0);
  const correct = records.reduce((sum, record) => sum + record.correctCount, 0);
  return {
    totalQuestions, practiced, attempts, correct,
    accuracy: attempts ? Math.round(correct / attempts * 100) : 0,
    wrong: records.filter(record => record.isWrong).length
  };
}

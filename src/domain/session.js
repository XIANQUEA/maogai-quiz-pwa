function shuffled(items, random) {
  const copy = [...items];
  for (let index = copy.length - 1; index > 0; index -= 1) {
    const swap = Math.floor(random() * (index + 1));
    [copy[index], copy[swap]] = [copy[swap], copy[index]];
  }
  return copy;
}

export function createSession(questions, options) {
  if (options.mode === "chapter") {
    const types = new Set(options.types || []);
    return questions.filter(question => question.chapterId === options.chapterId &&
      (types.size === 0 || types.has(question.type)));
  }
  if (options.mode === "wrong") {
    const ids = new Set(options.wrongIds || []);
    return questions.filter(question => ids.has(question.id));
  }
  if (options.mode === "random") {
    const count = Math.max(0, Number(options.count) || 0);
    return shuffled(questions, options.random || Math.random).slice(0, Math.min(count, questions.length));
  }
  throw new Error(`未知练习模式: ${options.mode}`);
}

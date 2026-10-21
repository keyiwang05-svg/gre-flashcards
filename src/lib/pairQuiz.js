const normalize = (value) => String(value ?? "").normalize("NFKC").trim().toLowerCase().replace(/\s+/g, " ");

export function pairQuizKey(pair) {
  const a = normalize(pair?.a), b = normalize(pair?.b);
  return a && b && a !== b ? JSON.stringify([a, b].sort()) : "";
}

function pairMap(pairs) {
  const result = new Map();
  for (const pair of pairs || []) {
    const key = pairQuizKey(pair);
    if (key && !result.has(key)) result.set(key, pair);
  }
  return result;
}

export function pairQuizPoolSignature(pairs) {
  return JSON.stringify([...pairMap(pairs).keys()].sort());
}

function shuffle(items, random) {
  const result = [...items];
  for (let i = result.length - 1; i > 0; i--) {
    const j = Math.floor(random() * (i + 1));
    [result[i], result[j]] = [result[j], result[i]];
  }
  return result;
}

function makeQuestion(target, allPairs, { random, getWordMeaning }, savedQuestion) {
  const terms = new Map(), synonyms = new Map();
  for (const pair of allPairs || []) {
    if (!pairQuizKey(pair)) continue;
    for (const [word, partner] of [[pair.a, pair.b], [pair.b, pair.a]]) {
      const key = normalize(word);
      if (!terms.has(key)) terms.set(key, word);
      if (!synonyms.has(key)) synonyms.set(key, new Set());
      synonyms.get(key).add(normalize(partner));
    }
  }

  const savedMatches = savedQuestion && pairQuizKey({ a: savedQuestion.pairA, b: savedQuestion.pairB }) === pairQuizKey(target)
    && [normalize(target.a), normalize(target.b)].includes(normalize(savedQuestion.correctAnswer))
    && normalize(savedQuestion.promptWord) !== normalize(savedQuestion.correctAnswer)
    && [normalize(target.a), normalize(target.b)].includes(normalize(savedQuestion.promptWord));
  const answerIsA = savedMatches ? normalize(savedQuestion.correctAnswer) === normalize(target.a) : random() > 0.5;
  const correctAnswer = answerIsA ? target.a : target.b;
  const promptWord = answerIsA ? target.b : target.a;
  const excluded = new Set([normalize(promptWord), normalize(correctAnswer), ...(synonyms.get(normalize(promptWord)) || [])]);
  const distractors = [...terms].filter(([key]) => !excluded.has(key)).map(([, word]) => word);
  const allowed = new Set([normalize(correctAnswer), ...distractors.map(normalize)]);
  const savedChoices = savedMatches && Array.isArray(savedQuestion.choices) ? savedQuestion.choices : [];
  const choicesAreValid = savedChoices.length > 0 && savedChoices.length <= 6
    && new Set(savedChoices.map(normalize)).size === savedChoices.length
    && savedChoices.includes(correctAnswer) && savedChoices.every((choice) => allowed.has(normalize(choice)));
  const choices = choicesAreValid ? [...savedChoices] : shuffle([correctAnswer, ...shuffle(distractors, random).slice(0, 5)], random);

  return {
    questionType: "bb_pairs", pairId: target.id, pairA: target.a, pairB: target.b,
    pairZh: target.zh || "", pairAMeaning: target.aMeaning || getWordMeaning(target.a) || "",
    pairBMeaning: target.bMeaning || getWordMeaning(target.b) || "", promptWord, correctAnswer, choices,
    promptZh: target.zh || "", explanation: `${target.a} = ${target.b}`,
    frequency: target.frequency || "", source: target.source || "",
  };
}

// The saved round contains only JSON data, so it can resume across reloads.
// Membership ignores statistics, array order, IDs, and A/B direction.
export function nextPairQuizQuestion(pool, allPairs, previous = null, options = {}) {
  const settings = { random: Math.random, getWordMeaning: () => "", resume: false, ...options };
  const pairs = pairMap(pool), keys = [...pairs.keys()];
  const signature = JSON.stringify([...keys].sort());
  if (!keys.length) return { question: null, round: null };

  const validQueue = previous?.signature === signature && Array.isArray(previous.queue)
    && previous.queue.length === keys.length && new Set(previous.queue).size === keys.length
    && previous.queue.every((key) => pairs.has(key))
    && Number.isInteger(previous.cursor) && previous.cursor >= 0 && previous.cursor <= keys.length;
  let round = validQueue ? { ...previous } : null;

  if (settings.resume && round?.cursor > 0 && round.question) {
    const target = pairs.get(round.queue[round.cursor - 1]);
    const question = makeQuestion(target, allPairs, settings, round.question);
    Object.assign(question, { roundNumber: round.number, roundPosition: round.cursor, roundTotal: keys.length });
    const answer = round.answer?.checked && question.choices.includes(round.answer.selectedChoice) ? round.answer : null;
    return { question, round: { ...round, question, answer } };
  }

  if (!round || round.cursor >= keys.length) {
    const queue = shuffle(keys, settings.random);
    const lastKey = previous?.lastKey;
    if (queue.length > 1 && queue[0] === lastKey) {
      const swapIndex = 1 + Math.floor(settings.random() * (queue.length - 1));
      [queue[0], queue[swapIndex]] = [queue[swapIndex], queue[0]];
    }
    const number = validQueue && Number.isSafeInteger(previous.number) && previous.number > 0 ? previous.number + 1 : 1;
    round = { signature, queue, cursor: 0, number };
  }

  const key = round.queue[round.cursor++];
  const question = makeQuestion(pairs.get(key), allPairs, settings);
  Object.assign(question, { roundNumber: round.number, roundPosition: round.cursor, roundTotal: keys.length });
  return { question, round: { ...round, lastKey: key, question, answer: null } };
}

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";
import { nextPairQuizQuestion, pairQuizKey, pairQuizPoolSignature } from "../src/lib/pairQuiz.js";

const pairs = JSON.parse(readFileSync(new URL("../public/data/builtinPairs.json", import.meta.url))).sixChoicePairs;
const serialize = (value) => JSON.parse(JSON.stringify(value));
function seededRandom(seed = 20261021) {
  return () => {
    seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
    return seed / 4294967296;
  };
}
const pair = (a, b, id = `${a}-${b}`) => ({ id, a, b, frequency: "超高频" });

test("all 1,014 builtin pairs appear once per round, with safe six-choice options", () => {
  const random = seededRandom(), adjacent = new Map();
  for (const p of pairs) for (const [a, b] of [[p.a, p.b], [p.b, p.a]]) {
    if (!adjacent.has(a)) adjacent.set(a, new Set());
    adjacent.get(a).add(b);
  }
  let round = null, lastKey;
  for (let number = 1; number <= 2; number++) {
    const seen = new Set();
    for (let i = 1; i <= pairs.length; i++) {
      const result = nextPairQuizQuestion(pairs, pairs, round, { random });
      round = result.round;
      const q = result.question, key = pairQuizKey({ a: q.pairA, b: q.pairB });
      assert.equal(q.roundNumber, number);
      assert.equal(q.roundPosition, i);
      assert.equal(q.roundTotal, 1014);
      assert.ok(!seen.has(key));
      if (i === 1 && lastKey) assert.notEqual(key, lastKey);
      seen.add(key);
      lastKey = key;
      assert.equal(q.choices.length, 6);
      assert.equal(new Set(q.choices.map((s) => s.toLowerCase())).size, 6);
      assert.ok(q.choices.includes(q.correctAnswer));
      assert.ok(!q.choices.includes(q.promptWord));
      for (const choice of q.choices) if (choice !== q.correctAnswer) {
        assert.ok(!adjacent.get(q.promptWord)?.has(choice), `${q.promptWord}: ${choice}`);
      }
    }
    assert.equal(seen.size, pairs.length);
  }
});

test("160-pair tier covers 100 distinct pairs in its first 100 questions", () => {
  const pool = pairs.filter((p) => p.frequency === "超高频"), random = seededRandom(17);
  let round = null;
  const ids = new Set();
  for (let i = 0; i < 100; i++) {
    const result = nextPairQuizQuestion(pool, pairs, serialize(round), { random });
    round = result.round;
    ids.add(result.question.pairId);
  }
  assert.equal(ids.size, 100);
  assert.equal(round.number, 1);
});

test("case, spaces, reversed pairs, and different IDs dedupe without losing other partners", () => {
  const pool = [pair(" Resolve ", "determination"), pair("DETERMINATION", "resolve", "other"), pair("resolve", "tenacity"), pair("Ａ", "Ｂ"), pair("B", "A")];
  let round = null;
  const seen = new Set();
  for (let i = 0; i < 3; i++) {
    const result = nextPairQuizQuestion(pool, pool, round);
    round = result.round;
    seen.add(pairQuizKey({ a: result.question.pairA, b: result.question.pairB }));
  }
  assert.equal(seen.size, 3);
  assert.equal(round.cursor, 3);
  assert.equal(nextPairQuizQuestion(pool, pool, round).question.roundNumber, 2);
});

test("resume preserves the question, choices, answer, and next position after JSON storage", () => {
  const first = nextPairQuizQuestion(pairs, pairs, null, { random: seededRandom(5) });
  first.round.answer = { checked: true, selectedChoice: first.question.correctAnswer };
  const restored = nextPairQuizQuestion(pairs, pairs, serialize(first.round), { resume: true, random: () => { throw Error("Resume must not reshuffle"); } });
  assert.deepEqual(restored.question, first.question);
  assert.deepEqual(restored.round.answer, first.round.answer);
  assert.equal(restored.round.cursor, 1);
  const next = nextPairQuizQuestion(pairs, pairs, restored.round);
  assert.equal(next.question.roundPosition, 2);
  assert.notEqual(next.question.pairId, first.question.pairId);
  assert.equal(next.round.answer, null);
});

test("updated statistics and array order do not restart or reshuffle a round", () => {
  const first = nextPairQuizQuestion(pairs, pairs);
  const changed = pairs.map((p) => ({ ...p, stats: { seen: 99, correct: 99 } })).reverse();
  assert.equal(pairQuizPoolSignature(changed), pairQuizPoolSignature(pairs));
  const next = nextPairQuizQuestion(changed, changed, first.round);
  assert.deepEqual(next.round.queue, first.round.queue);
  assert.equal(next.round.cursor, 2);
});

test("known synonyms outside the selected tier or search never become distractors", () => {
  const target = pair("resolve", "determination");
  const bank = [target, pair("resolve", "tenacity"), ...Array.from({ length: 6 }, (_, i) => pair(`unrelated${i}`, `other${i}`))];
  const q = nextPairQuizQuestion([target], bank, null, { random: () => 0.2 }).question;
  assert.equal(q.promptWord, "resolve");
  assert.equal(q.correctAnswer, "determination");
  assert.ok(!q.choices.includes("tenacity"));
  assert.equal(q.choices.length, 6);
});

test("synonym exclusion uses direct pairs without merging different senses transitively", () => {
  const target = pair("a", "b"), bank = [target, pair("b", "c"), pair("d", "e")];
  const q = nextPairQuizQuestion([target], bank, null, { random: () => 0.2 }).question;
  assert.equal(q.promptWord, "a");
  assert.ok(q.choices.includes("c"));
});

test("empty filters do not silently fall back to the full library", () => {
  assert.deepEqual(nextPairQuizQuestion([], pairs), { question: null, round: null });
  assert.deepEqual(nextPairQuizQuestion([pair("same", "same"), pair("", "a")], pairs), { question: null, round: null });
});

test("changed pool and malformed saved queues start a valid new round", () => {
  const first = nextPairQuizQuestion(pairs, pairs);
  const subset = pairs.slice(0, 3);
  for (const invalid of [first.round, { ...first.round, cursor: -1 }, { ...first.round, queue: ["missing"] }, { signature: pairQuizPoolSignature(subset), queue: Array(3).fill(pairQuizKey(subset[0])), cursor: 0 }]) {
    const result = nextPairQuizQuestion(subset, pairs, invalid);
    assert.equal(result.round.cursor, 1);
    assert.equal(new Set(result.round.queue).size, 3);
    assert.equal(result.question.roundTotal, 3);
  }
});

test("a single-pair pool works across rounds and uses all-bank word meanings", () => {
  const p = pair("a", "b"), getWordMeaning = (s) => `meaning ${s}`;
  const first = nextPairQuizQuestion([p], [p], null, { getWordMeaning });
  const second = nextPairQuizQuestion([p], [p], first.round, { getWordMeaning });
  assert.equal(second.question.roundNumber, 2);
  assert.deepEqual(new Set(second.question.choices), new Set([second.question.correctAnswer]));
  assert.equal(second.question.pairAMeaning, "meaning a");
  assert.equal(second.question.pairBMeaning, "meaning b");
});

test("the last question resumes without starting the next round until Next is pressed", () => {
  const pool = pairs.slice(0, 2);
  const first = nextPairQuizQuestion(pool, pairs);
  const last = nextPairQuizQuestion(pool, pairs, first.round);
  const resumed = nextPairQuizQuestion(pool, pairs, serialize(last.round), { resume: true });
  assert.deepEqual(resumed.question, last.question);
  assert.equal(resumed.question.roundNumber, 1);
  assert.equal(nextPairQuizQuestion(pool, pairs, resumed.round).question.roundNumber, 2);
});

test("new synonym relationships also invalidate previously cached distractors on resume", () => {
  const p = pair("a", "b"), bank = [p, pair("c", "d")];
  const first = nextPairQuizQuestion([p], bank, null, { random: () => 0.2 });
  assert.ok(first.question.choices.includes("c"));
  const resumed = nextPairQuizQuestion([p], [...bank, pair("a", "c")], first.round, { resume: true });
  assert.ok(!resumed.question.choices.includes("c"));
  assert.equal(resumed.round.cursor, 1);
});

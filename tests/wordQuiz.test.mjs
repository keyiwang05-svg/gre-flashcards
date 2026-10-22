import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";
import test from "node:test";

// Exercise the exact question builder currently wired into the app.
const app = readFileSync(new URL("../src/GREFlashcardsApp.tsx", import.meta.url), "utf8");
const start = app.indexOf("function buildQuizQuestion(");
const end = app.indexOf("\nasync function loadBuiltinLibraries", start);
assert.ok(start >= 0 && end > start, "Quiz builder must be found in the app");
const source = app.slice(start, end);
const context = {
  pickRandom: (items) => items[0],
  uniq: (items) => [...new Set(items)],
  shuffleArray: (items) => [...items],
};
const buildQuizQuestion = vm.runInNewContext(source + "\nbuildQuizQuestion", context);
const word = (name, synonyms = [], antonyms = []) => ({
  id: name, word: name, senses: [{ zh: name, synonyms, antonyms }],
});
const words = [
  word("all-important", ["negligible"], ["negligible"]),
  word("irrelevant"),
  word("ordinary"),
  word("unrelated"),
  word("different"),
  word("minor"),
];
const bbPairs = [
  { a: "all-important", b: "crucial" },
  { a: "all-important", b: "vital" },
];

test("equivalence never treats erroneous synonyms as correct", () => {
  const q = buildQuizQuestion(words, "equivalence", bbPairs);
  assert.equal(q.correctAnswer, "crucial");
  assert.ok(!q.choices.includes("negligible") || q.correctAnswer !== "negligible");
  assert.ok(!q.choices.includes("vital"), "another approved partner must not become a distractor");
});

test("equivalence has no question when no verified BB partner exists", () => {
  assert.equal(buildQuizQuestion(words, "equivalence", []), null);
});

test("correct answer remains verified under reverse pair orientation", () => {
  const q = buildQuizQuestion(words, "equivalence", [{ a: "crucial", b: "all-important" }]);
  assert.equal(q.correctAnswer, "crucial");
});

test("antonym mode is separate from equivalence mode", () => {
  const q = buildQuizQuestion(words, "antonym", bbPairs);
  assert.equal(q.correctAnswer, "negligible");
  assert.equal(q.questionType, "antonym");
});

test("equivalence choices are distinct and omit every known correct partner other than the answer", () => {
  const q = buildQuizQuestion(words, "equivalence", bbPairs);
  assert.equal(q.choices.length, 4);
  assert.equal(new Set(q.choices.map((x) => x.toLowerCase())).size, 4);
  assert.ok(q.choices.includes("crucial"));
  assert.ok(!q.choices.includes("vital"));
});

test("existing 1014 BB pairs are unchanged and still the sole fixed pair library", () => {
  const pairs = JSON.parse(readFileSync(new URL("../public/data/builtinPairs.json", import.meta.url), "utf8")).sixChoicePairs;
  assert.equal(pairs.length, 1014);
});

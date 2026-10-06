import test from "node:test";
import assert from "node:assert/strict";
import { parsePageList, parseRangeGroups, chunkBySize, formatPageList } from "../server/lib/ranges.js";

test("parsePageList 支持区间、单页、开区间与 all", () => {
  assert.deepEqual(parsePageList("1-3,5", 10), [0, 1, 2, 4]);
  assert.deepEqual(parsePageList("8-", 10), [7, 8, 9]);
  assert.deepEqual(parsePageList("all", 3), [0, 1, 2]);
  assert.deepEqual(parsePageList("", 3), [0, 1, 2]);
  assert.deepEqual(parsePageList("3,1,2,2", 5), [0, 1, 2]);
  assert.deepEqual(parsePageList("20-30", 5), []);
  assert.throws(() => parsePageList("abc", 5));
});

test("parseRangeGroups 保留分组", () => {
  assert.deepEqual(parseRangeGroups("1-2,4,6-", 7), [[0, 1], [3], [5, 6]]);
  assert.throws(() => parseRangeGroups("", 5));
});

test("chunkBySize / formatPageList", () => {
  assert.deepEqual(chunkBySize([0, 1, 2, 3, 4], 2), [[0, 1], [2, 3], [4]]);
  assert.equal(formatPageList([0, 1, 2, 4]), "1-3,5");
  assert.equal(formatPageList([5]), "6");
});

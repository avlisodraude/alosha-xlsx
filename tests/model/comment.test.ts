import { beforeEach, describe, expect, it } from "vitest";
import { type Comment, Comments } from "../../src/model/comment.js";
import { Workbook } from "../../src/model/workbook.js";
import type { Worksheet } from "../../src/model/worksheet.js";

const note: Comment = { texts: "hello", author: "Alice" };

describe("Comments manager", () => {
  let comments: Comments;

  beforeEach(() => {
    comments = new Comments();
  });

  it("add / find by exact ref", () => {
    comments.add("B2", note);
    expect(comments.find("B2")).toEqual(note);
    expect(comments.find("B3")).toBeUndefined();
    expect(comments.size).toBe(1);
  });

  it("remove and clear", () => {
    comments.add("A1", note);
    comments.remove("A1");
    expect(comments.find("A1")).toBeUndefined();
    comments.add("A1", note);
    comments.clear();
    expect(comments.size).toBe(0);
  });

  it("round-trips through the model in insertion order", () => {
    comments.add("A1", { texts: "one" });
    comments.add("B2", { texts: "two", author: "Bob" });
    const model = comments.model;
    expect(model).toEqual([
      { ref: "A1", texts: "one" },
      { ref: "B2", texts: "two", author: "Bob" },
    ]);

    const restored = new Comments();
    restored.model = model;
    expect(restored.find("A1")).toEqual({ texts: "one" });
    expect(restored.find("B2")).toEqual({ texts: "two", author: "Bob" });
  });
});

describe("Cell.note accessor", () => {
  let ws: Worksheet;

  beforeEach(() => {
    ws = new Workbook().addWorksheet("Data");
  });

  it("routes get/set through the worksheet's comments store", () => {
    ws.getCell("B2").note = { texts: "hi", author: "Alice" };
    expect(ws.comments.find("B2")).toEqual({ texts: "hi", author: "Alice" });
    expect(ws.getCell("B2").note).toEqual({ texts: "hi", author: "Alice" });
    expect(ws.getCell("C3").note).toBeUndefined();
  });

  it("preserves an unset author (the default is applied only at write time)", () => {
    ws.getCell("A1").note = { texts: "no author" };
    expect(ws.getCell("A1").note?.author).toBeUndefined();
  });
});

describe("WorksheetModel.comments round-trip", () => {
  it("serializes comments only when present and reloads them", () => {
    const ws = new Workbook().addWorksheet("Data");
    expect(ws.model.comments).toBeUndefined();

    ws.getCell("B2").note = { texts: "hi", author: "Alice" };
    const model = ws.model;
    expect(model.comments).toEqual([{ ref: "B2", texts: "hi", author: "Alice" }]);

    const reloaded = new Workbook().addWorksheet("Other");
    reloaded.model = model;
    expect(reloaded.getCell("B2").note).toEqual({ texts: "hi", author: "Alice" });
  });
});

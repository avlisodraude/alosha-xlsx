import { describe, expect, it } from "vitest";
import { dateToSerial, serialToDate } from "../../src/xlsx/date.js";

const utc = (y: number, m: number, d: number) => new Date(Date.UTC(y, m - 1, d));

describe("dateToSerial / serialToDate (1900 system)", () => {
  it("anchors serial 1 to 1900-01-01", () => {
    expect(dateToSerial(utc(1900, 1, 1))).toBe(1);
    expect(serialToDate(1).getTime()).toBe(utc(1900, 1, 1).getTime());
  });

  it("anchors serial 2 to 1900-01-02", () => {
    expect(dateToSerial(utc(1900, 1, 2))).toBe(2);
    expect(serialToDate(2).getTime()).toBe(utc(1900, 1, 2).getTime());
  });

  it("maps serial 60 to the fictional 1900-02-29 (normalized to 1900-03-01)", () => {
    // 1900-02-29 does not exist; JS normalizes it to 1900-03-01.
    expect(serialToDate(60).getTime()).toBe(utc(1900, 2, 29).getTime());
    expect(serialToDate(60).getTime()).toBe(utc(1900, 3, 1).getTime());
  });

  it("shifts real dates across the leap-bug boundary", () => {
    expect(dateToSerial(utc(1900, 2, 28))).toBe(59);
    // No real date maps to serial 60; 1900-03-01 is 61.
    expect(dateToSerial(utc(1900, 3, 1))).toBe(61);
  });

  it("matches the Unix-epoch anchor 25569 = 1970-01-01", () => {
    expect(dateToSerial(utc(1970, 1, 1))).toBe(25569);
    expect(serialToDate(25569).getTime()).toBe(utc(1970, 1, 1).getTime());
  });

  it("matches a modern anchor 44197 = 2021-01-01", () => {
    expect(dateToSerial(utc(2021, 1, 1))).toBe(44197);
    expect(serialToDate(44197).getTime()).toBe(utc(2021, 1, 1).getTime());
  });

  it("carries time-of-day in the fractional part", () => {
    const noon = new Date(Date.UTC(2021, 0, 1, 12, 0, 0));
    expect(dateToSerial(noon)).toBe(44197.5);
    expect(serialToDate(44197.5).getTime()).toBe(noon.getTime());
  });
});

describe("dateToSerial / serialToDate (1904 system)", () => {
  it("anchors serial 0 to 1904-01-01", () => {
    expect(dateToSerial(utc(1904, 1, 1), true)).toBe(0);
    expect(serialToDate(0, true).getTime()).toBe(utc(1904, 1, 1).getTime());
  });

  it("is offset from the 1900 system by 1462 days", () => {
    expect(dateToSerial(utc(1904, 1, 1), false)).toBe(1462);
    const sample = utc(2021, 6, 15);
    expect(dateToSerial(sample, false) - dateToSerial(sample, true)).toBe(1462);
  });
});

describe("round-trips", () => {
  const dates = [
    utc(1900, 1, 1),
    utc(1900, 2, 28),
    utc(1900, 3, 1),
    utc(1970, 1, 1),
    utc(2021, 1, 1),
    utc(2024, 2, 29),
    new Date(Date.UTC(2024, 5, 15, 8, 30, 0)),
  ];

  for (const d of dates) {
    it(`round-trips ${d.toISOString()} in both systems`, () => {
      expect(serialToDate(dateToSerial(d)).getTime()).toBe(d.getTime());
      expect(serialToDate(dateToSerial(d, true), true).getTime()).toBe(d.getTime());
    });
  }
});

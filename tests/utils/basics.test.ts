import { describe, expect, it } from "vitest";

import { currentTime } from "../../src/utils/basics";

describe("currentTime", () => {
  it("returns a timestamp in the displayed UK locale format", () => {
    expect(currentTime()).toMatch(/^\d{2}\/\d{2}\/\d{2}, \d{2}:\d{2}$/);
  });
});

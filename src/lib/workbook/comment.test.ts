import { describe, expect, it } from "vitest";
import {
  addCommentReply,
  createCommentThread,
  normalizeCellComment,
  setCommentResolved,
} from "@/lib/workbook/comment";

describe("comment threads", () => {
  it("migrates legacy single comment", () => {
    const legacy = {
      id: "1",
      userId: "u",
      authorName: "Ann",
      text: "Hello",
      createdAt: "2020-01-01",
      updatedAt: "2020-01-01",
    };
    const thread = normalizeCellComment(legacy)!;
    expect(thread.messages).toHaveLength(1);
    expect(thread.messages[0]?.text).toBe("Hello");
  });

  it("supports replies and resolve", () => {
    let thread = createCommentThread("u1", "Ann", "First");
    thread = addCommentReply(thread, "u2", "Bob", "Reply");
    expect(thread.messages).toHaveLength(2);
    thread = setCommentResolved(thread, true);
    expect(thread.status).toBe("resolved");
  });
});
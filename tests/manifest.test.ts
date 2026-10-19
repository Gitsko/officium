import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

const manifest = JSON.parse(readFileSync(new URL("../manifest.json", import.meta.url), "utf8"));

describe("manifest", () => {
  it("carries the Officium identity instead of the template defaults", () => {
    expect(manifest.name.short).toBe("Officium");
    expect(manifest.id).not.toBe("f2b103f1-1ab1-4e1b-8f0b-072aa3d4e19d");
    expect(JSON.stringify(manifest)).not.toMatch(/contoso/i);
  });

  // New Outlook has no item-less calendar surface, so the pane hangs off opened items.
  it("offers the task pane on read mails and on appointments", () => {
    const contexts = manifest.extensions[0].ribbons.flatMap((r: { contexts: string[] }) => r.contexts);
    expect(contexts).toEqual(
      expect.arrayContaining(["mailRead", "meetingDetailsAttendee", "meetingDetailsOrganizer"])
    );
  });

  // Calendar data comes from Graph, so the add-in never needs to read the opened item.
  it("requests only restricted mailbox access", () => {
    const names = manifest.authorization.permissions.resourceSpecific.map((p: { name: string }) => p.name);
    expect(names).toEqual(["MailboxItem.Restricted.User"]);
  });

  it("declares German as the default language, matching the task pane", () => {
    expect(manifest.localizationInfo.defaultLanguageTag).toBe("de-de");
  });

  it("only references actions that a runtime defines", () => {
    const ext = manifest.extensions[0];
    const actionIds = ext.runtimes.flatMap((r: { actions: { id: string }[] }) => r.actions.map((a) => a.id));
    const referenced = JSON.stringify(ext.ribbons).match(/"actionId":"[^"]+"/g) ?? [];
    for (const ref of referenced) {
      expect(actionIds).toContain(ref.slice('"actionId":"'.length, -1));
    }
  });
});

import { describe, expect, it } from "vitest";
import { releaseNotes } from "../scripts/release-notes.mjs";

const CHANGELOG = [
  "# Changelog",
  "",
  "## [Unreleased]",
  "",
  "## [0.2.0] - 2026-09-13",
  "",
  "### Changed (breaking)",
  "",
  "- `mail send` saves a draft.",
  "",
  "### Fixed",
  "",
  "- Multi-line bodies survive.",
  "",
  "## [0.1.1] - 2026-09-13",
  "",
  "### Fixed",
  "",
  "- Older fix.",
].join("\n");

describe("release notes", () => {
  it("extracts the section for a tag, including its subsections", () => {
    const notes = releaseNotes(CHANGELOG, "v0.2.0");
    expect(notes).toContain("### Changed (breaking)");
    expect(notes).toContain("- `mail send` saves a draft.");
    expect(notes).toContain("### Fixed");
    expect(notes).toContain("- Multi-line bodies survive.");
  });

  it("stops at the next version heading", () => {
    expect(releaseNotes(CHANGELOG, "0.2.0")).not.toContain("Older fix.");
    expect(releaseNotes(CHANGELOG, "0.1.1")).not.toContain("Multi-line bodies survive.");
  });

  it("falls back when the version has no section yet", () => {
    expect(releaseNotes(CHANGELOG, "v9.9.9")).toBe("See CHANGELOG.md.");
    expect(releaseNotes(CHANGELOG, "v9.9.9", "custom fallback")).toBe("custom fallback");
  });

  it("falls back for an empty section", () => {
    expect(releaseNotes("## [Unreleased]\n\n## [0.2.0]\n", "0.2.0")).toBe(
      "See CHANGELOG.md.",
    );
  });
});

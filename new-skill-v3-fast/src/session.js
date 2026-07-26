// session.json is the orchestration state (the "SESSION.JSON" box in the diagram):
// deterministic ordering, per-page status, timing, retries, and audit trail.
import { readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";

export const STATUS = {
  PENDING: "pending",
  IN_PROGRESS: "in_progress",
  COMPLETED: "completed",
  FAILED: "failed",
};

export function sessionPath(workDir) {
  return join(workDir, "session.json");
}

export function loadSession(workDir) {
  return JSON.parse(readFileSync(sessionPath(workDir), "utf8"));
}

export function saveSession(workDir, session) {
  session.updatedAt = new Date().toISOString();
  writeFileSync(sessionPath(workDir), JSON.stringify(session, null, 2));
}

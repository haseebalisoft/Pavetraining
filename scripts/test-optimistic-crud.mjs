#!/usr/bin/env node
/**
 * Offline behavioural test for the optimistic-UI logic added to
 * AdminCrudPage (companies delete). Replicates the client-side state
 * transitions so we can validate:
 *
 *   1. Delete success                  → row gone, no review entry.
 *   2. Delete API fails                → row silently restored, review entry added.
 *   3. Delete of unknown id            → no-op, no crash.
 *   4. Two concurrent deletes, one ok  → only the failing one is restored.
 *   5. Slow network / late failure     → other actions still work meanwhile.
 *   6. Retry after failure             → row disappears again once retried and succeeds.
 *
 * Uses a tiny fake state machine that mirrors the real component's mutation
 * pattern (setRows, setSelectedIds, setReviewOps).
 *
 * Usage:  node scripts/test-optimistic-crud.mjs
 */

const green = (s) => `\x1b[32m${s}\x1b[0m`;
const red = (s) => `\x1b[31m${s}\x1b[0m`;
const bold = (s) => `\x1b[1m${s}\x1b[0m`;
const dim = (s) => `\x1b[2m${s}\x1b[0m`;

let checks = 0;
let failed = 0;
function assert(label, got, expected) {
  checks += 1;
  const ok = JSON.stringify(got) === JSON.stringify(expected);
  if (!ok) failed += 1;
  console.log(
    `    ${ok ? green("PASS") : red("FAIL")}  ${label}${
      ok ? "" : dim(`\n         expected ${JSON.stringify(expected)}\n         got      ${JSON.stringify(got)}`)
    }`,
  );
}

/** Simulates the AdminCrudPage optimistic delete state machine. */
function createOptimisticStore(initialRows) {
  const state = {
    rows: [...initialRows],
    selectedIds: new Set(),
    reviewOps: [],
  };

  function labelFor(row) {
    return row.companyName || `#${row.id}`;
  }

  async function deleteOne(id, apiFn) {
    const snapshot = state.rows.find((row) => row.id === id);
    if (!snapshot) return;
    // Optimistic: remove immediately.
    state.selectedIds.delete(id);
    state.rows = state.rows.filter((row) => row.id !== id);
    try {
      await apiFn();
    } catch (error) {
      // Silent restore.
      if (!state.rows.some((row) => row.id === id)) {
        state.rows = [...state.rows, snapshot];
      }
      state.reviewOps.push({
        key: `delete-${id}-${Date.now()}-${Math.random()}`,
        op: "delete",
        label: labelFor(snapshot),
        errorMessage: error instanceof Error ? error.message : String(error),
      });
    }
  }

  return { state, deleteOne };
}

async function scenario(title, fn) {
  console.log(`\n${bold(title)}`);
  await fn();
}

const rowsBase = [
  { id: "1", companyName: "Alpha" },
  { id: "2", companyName: "Beta" },
  { id: "3", companyName: "Gamma" },
];

await scenario("1. Delete success removes the row and does not log a review entry", async () => {
  const { state, deleteOne } = createOptimisticStore(rowsBase);
  await deleteOne("2", async () => {
    /* success */
  });
  assert("row hidden immediately", state.rows.map((r) => r.id), ["1", "3"]);
  assert("no review entry logged", state.reviewOps.length, 0);
});

await scenario("2. Delete API failure silently restores the row + logs review", async () => {
  const { state, deleteOne } = createOptimisticStore(rowsBase);
  await deleteOne("2", async () => {
    throw new Error("SharePoint refused: related row exists");
  });
  assert("row restored", state.rows.map((r) => r.id).sort(), ["1", "2", "3"]);
  assert("one review entry", state.reviewOps.length, 1);
  assert("review label = Beta", state.reviewOps[0].label, "Beta");
  assert(
    "review error message preserved",
    state.reviewOps[0].errorMessage,
    "SharePoint refused: related row exists",
  );
});

await scenario("3. Delete of unknown id is a no-op (no crash, no state change)", async () => {
  const { state, deleteOne } = createOptimisticStore(rowsBase);
  await deleteOne("999", async () => {
    throw new Error("should not be called");
  });
  assert("rows unchanged", state.rows.map((r) => r.id), ["1", "2", "3"]);
  assert("no review entry", state.reviewOps.length, 0);
});

await scenario("4. Two concurrent deletes — one succeeds, one fails", async () => {
  const { state, deleteOne } = createOptimisticStore(rowsBase);
  await Promise.all([
    deleteOne("1", async () => {
      /* success */
    }),
    deleteOne("2", async () => {
      throw new Error("network flake");
    }),
  ]);
  assert("only failing row restored", state.rows.map((r) => r.id).sort(), ["2", "3"]);
  assert("exactly one review entry", state.reviewOps.length, 1);
  assert("failing entry is Beta", state.reviewOps[0].label, "Beta");
});

await scenario("5. Slow delete → other action fires meanwhile → both settle correctly", async () => {
  const { state, deleteOne } = createOptimisticStore(rowsBase);
  const slow = deleteOne("1", async () => {
    await new Promise((r) => setTimeout(r, 60));
    throw new Error("slow-failure");
  });
  const fast = deleteOne("3", async () => {
    /* success */
  });
  await Promise.all([slow, fast]);
  assert("Alpha restored, Gamma gone", state.rows.map((r) => r.id).sort(), [
    "1",
    "2",
  ]);
  assert("one review entry (Alpha)", state.reviewOps.map((r) => r.label), [
    "Alpha",
  ]);
});

// --- Bulk-delete simulation (matches deleteSelected optimistic path) ---
function createBulkOptimisticStore(initialRows) {
  const state = {
    rows: [...initialRows],
    hiddenIds: new Set(),
    reviewOps: [],
  };

  function label(count) {
    return `${count} record${count === 1 ? "" : "s"}`;
  }

  async function bulkDelete(ids, apiFn) {
    const snapshotById = new Map();
    for (const id of ids) {
      const row = state.rows.find((r) => r.id === id);
      if (row) snapshotById.set(id, row);
    }
    const targetIds = [...snapshotById.keys()];
    if (targetIds.length === 0) return;
    // Optimistic: hide all + persist hidden.
    const target = new Set(targetIds);
    state.rows = state.rows.filter((row) => !target.has(row.id));
    for (const id of targetIds) state.hiddenIds.add(id);
    try {
      await apiFn();
      // "load()" reconciliation would drop hiddenIds not seen in payload;
      // simulate: everything that got hidden also got deleted for real.
      for (const id of targetIds) state.hiddenIds.delete(id);
    } catch (error) {
      // Restore all rows silently.
      const known = new Set(state.rows.map((row) => row.id));
      for (const id of targetIds) {
        if (!known.has(id)) state.rows.push(snapshotById.get(id));
        state.hiddenIds.delete(id);
      }
      state.reviewOps.push({
        op: "delete",
        label: label(targetIds.length),
        errorMessage:
          error instanceof Error ? error.message : String(error),
      });
    }
  }

  return { state, bulkDelete };
}

await scenario(
  "7. Bulk delete SUCCESS hides every selected row and persists in hiddenIds until reconciled",
  async () => {
    const { state, bulkDelete } = createBulkOptimisticStore(rowsBase);
    let seenHidden = 0;
    await bulkDelete(["1", "2"], async () => {
      // At mid-flight, snapshot the state.
      seenHidden = state.hiddenIds.size;
      await new Promise((r) => setTimeout(r, 10));
    });
    assert("rows remaining after bulk delete", state.rows.map((r) => r.id), ["3"]);
    assert("mid-flight hiddenIds size = 2", seenHidden, 2);
    assert("post-success hiddenIds cleared", state.hiddenIds.size, 0);
    assert("no review entry logged", state.reviewOps.length, 0);
  },
);

await scenario(
  "8. Bulk delete FAILURE restores every row + logs one review entry, keeps others gone if selective",
  async () => {
    const { state, bulkDelete } = createBulkOptimisticStore(rowsBase);
    await bulkDelete(["1", "2"], async () => {
      throw new Error("Batch failed");
    });
    assert(
      "rows fully restored",
      state.rows.map((r) => r.id).sort(),
      ["1", "2", "3"],
    );
    assert("one review entry", state.reviewOps.length, 1);
    assert(
      "review label is combined count",
      state.reviewOps[0].label,
      "2 records",
    );
    assert("hidden queue cleared", state.hiddenIds.size, 0);
  },
);

await scenario(
  "9. Bulk delete then a page refresh hides those rows on SSR too",
  async () => {
    const { state, bulkDelete } = createBulkOptimisticStore(rowsBase);
    // Fire and DON'T await — simulate mid-cascade refresh.
    const pending = bulkDelete(["1", "3"], async () => {
      await new Promise((r) => setTimeout(r, 40));
    });
    // "Refresh" arrives while cascade is still running.
    await new Promise((r) => setTimeout(r, 5));
    // SSR would return all 3 rows fresh from SharePoint; the client filters
    // them against hiddenIds.
    const ssrPayload = [...rowsBase];
    const visibleAfterRefresh = ssrPayload.filter(
      (row) => !state.hiddenIds.has(row.id),
    );
    assert(
      "SSR-derived view respects hiddenIds during cascade",
      visibleAfterRefresh.map((r) => r.id),
      ["2"],
    );
    await pending;
  },
);

await scenario("6. Retry after failure eventually succeeds", async () => {
  const { state, deleteOne } = createOptimisticStore(rowsBase);
  await deleteOne("2", async () => {
    throw new Error("transient");
  });
  assert("row initially restored", state.rows.map((r) => r.id).sort(), [
    "1",
    "2",
    "3",
  ]);
  assert("review entry present", state.reviewOps.length, 1);
  // Retry (dismiss the entry then delete again with success).
  state.reviewOps = [];
  await deleteOne("2", async () => {
    /* success on retry */
  });
  assert("row gone after retry", state.rows.map((r) => r.id).sort(), ["1", "3"]);
  assert("no review entry left", state.reviewOps.length, 0);
});

console.log(
  "\n" +
    (failed === 0
      ? green(bold(`ALL GOOD — ${checks} checks passed.`))
      : red(bold(`${failed}/${checks} checks failed.`))),
);
process.exit(failed === 0 ? 0 : 1);

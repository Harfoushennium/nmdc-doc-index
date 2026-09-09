title: NMDC Document Index — Cycle 1 Source Profiler

# Repository Task

Collaboration-ID: `NMDC-DOC-INDEX-001`
Repository: `Harfoushennium/nmdc-doc-index`
Authorized cycle: `1`
Reviewer: `ChatGPT Browser`

Execute **Cycle 1 only** for this repository.

The authoritative task specification is `CYCLE_1_ASSIGNMENT.md`. Read and follow it completely before making any implementation change.

Also read, in this order:

1. `README.md`
2. `PROJECT_SPEC.md`
3. `CLASSIFICATION_MODEL.md`
4. `AGENTS.md`
5. `CYCLE_1_ASSIGNMENT.md`
6. the complete current Agent Collaboration PR discussion

If any instruction conflicts, stop and report the conflict instead of choosing silently.

## Hard gate

This authorization is limited to **read-only source profiling and classification discovery**.

Do not:

- modify anything under `DATA/`;
- build the final consolidated index;
- create `NMDC_DOCUMENT_INDEX.xlsx`;
- build the full Document → Revision → Event extraction pipeline;
- silently approve Classification Model v2;
- begin Cycle 2;
- merge;
- force-push.

Unknown or ambiguous classifications must remain visible as `UNCLASSIFIED` / `REVIEW_REQUIRED`.

After completing the Cycle-1 deliverables, tests, and real-data profiling required by `CYCLE_1_ASSIGNMENT.md`, post exactly one truthful `HERMES_REPORT` to the draft PR and stop.

Cycle 2 is not authorized until the user and independent ChatGPT reviewer approve Classification Model v2.
# Version 2 Deployment Readiness

## Completed
- [x] Record-aware comparison engine
- [x] Row movement does not count as a regulatory change
- [x] 47 tracked BPU standards configured
- [x] GO / GOP / TO / DP profile
- [x] NERC-first Control Room interface
- [x] Separate FERC Watch view
- [x] Morning Brief and priority counts
- [x] Standard-family and individual-standard views
- [x] Search and filtering
- [x] Health/status indicators
- [x] Daily workflow definition
- [x] Unit tests for semantic comparison
- [x] Data-contract validation
- [x] Python and JavaScript syntax validation in GitHub Actions
- [x] Public-source-only architecture
- [x] Write core V2 files to `v2-unified-control-room`
- [x] V2 branch validation workflow passes in GitHub Actions

## Before production merge
- [ ] Run tracker workflow manually on `v2-unified-control-room`
- [ ] Confirm first run creates semantic baseline
- [ ] Run a second scan and verify no row-shift false positives
- [ ] Verify Gmail delivery to configured recipients
- [ ] Inspect generated V2 dashboard data and standard-family counts
- [ ] Verify GitHub Pages rendering after production merge
- [ ] Review any genuine tracked-standard changes
- [ ] Reconcile current `main` data commits with V2 branch
- [ ] Merge once into production

No confidential registry, evidence, or operational information belongs in the public site.

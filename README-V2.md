# BPU NERC Regulatory Control Room — Version 2 Release Candidate

This build is designed to run at **$0 additional cost** using GitHub Actions, GitHub Pages, Gmail SMTP, Python, HTML/CSS/JavaScript, and JSON.

## Major Version 2 changes

- Replaces cell-address comparison with a semantic record comparison engine.
- Row movement in the NERC workbook no longer creates thousands of false-positive changes.
- Establishes a fresh semantic baseline on the first Version 2 run.
- Focuses the dashboard and email on BPU's tracked standards.
- Registered-function profile: GO, GOP, TO, DP.
- NERC remains the default Control Room view; FERC remains separate.
- Adds standards matrix, family filters, priority classification, search, scan history, and health indicators.
- Leaves confidential registry, evidence, and operational information outside the public site.

## Important

The supplied screenshots contain **47 tracked standards**. An earlier conversation count of 42 was a miscount; the configuration in this build uses all 47 standards visible in the screenshots.

## First run after deployment

The first run intentionally creates `state/nerc-records.json` and reports **Semantic baseline established**. The second and later runs perform record-aware comparison against that baseline.

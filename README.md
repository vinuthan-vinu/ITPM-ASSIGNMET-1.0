# ITPM Assignment 1.0 - Option 2

This repository contains the Playwright automation and test documentation for:

Functional and usability testing of https://www.pixelssuite.com/

## Contents

- `tests/resize-preview.spec.js` - Playwright automated test for Resize Image preview.
- `fixtures/sample-preview.png` - PNG file used by the automated test.
- `execution_results.csv` - Recorded execution result for the automated scenario.
- `Manual Test Cases for Option 2.xlsx` - Manual test cases for the remaining 35 scenarios.
- `repository_link.txt` - Public GitHub repository link.

## Automated Scenario

The automated scenario verifies that the Resize Image feature displays a valid PNG in the Preview area after upload.

Test case ID: `Pos_0001`

## Setup

Install dependencies:

```bash
npm install
npx playwright install
```

## Run Tests

```bash
npm test
```

The test writes the latest result to `execution_results.csv` and saves screenshot evidence in `test-results/resize-preview.png`.

## Notes

- The test targets the live site at `https://www.pixelssuite.com/resize-image`.
- The manual Excel file records the remaining 35 scenarios required by the assignment brief.

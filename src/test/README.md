# Testing Infrastructure

This directory contains the automated tests for the Instant Issues List application.

## Test Framework

We use **Vitest** for testing, which provides:
- Fast test execution with native ES modules support
- Jest-compatible API for familiar testing patterns
- Built-in TypeScript support
- Integration with Vite for optimal performance

## Test Structure

- `setup.ts` - Test environment setup and global configurations
- `*.test.ts` - Unit tests for utility functions and types
- `*.test.tsx` - Component tests using React Testing Library

## Current Test Coverage

### Utility Functions (`utils.test.ts`)
- `dateToday()` - Date formatting functionality
- `formatCommentDate()` - Comment date parsing and formatting

### Type Definitions (`types.test.ts`)
- Interface and type validations
- Custom class implementations (Break)

### React Components (`App.test.tsx`)
- Basic component rendering
- Essential UI element presence

### Automatic Numbering (`numberingUtils.test.ts`)
- Number formatting functions (toLowerLetter, toUpperLetter, toLowerRoman, toUpperRoman)
- Ordinal and cardinal text conversion (toOrdinal, toCardinalText, toOrdinalText)
- Number formatting dispatcher (formatNumber)
- Counter management (initializeCounters, updateCounters)
- Document parsing helpers (buildNumberingMaps, buildStyleMaps)
- Style hierarchy resolution (resolveStyleNumbering, extractParagraphStyle)
- Edge cases and error handling for all numbering functions

### Extraction Numbering (`wordExtraction.test.ts`)
- Carry-forward of the last rendered numbering string for later interesting paragraphs
- Style-linked numbering resolved from `styles.xml` and `numbering.xml`
- Fallback to section/page references when numbering metadata is absent
- Target-behavior fixtures for restart and override handling (`it.fails`) to guide future OOXML parser work

## Running Tests

```bash
# Run tests in watch mode (development)
npm run test

# Run tests once (CI/production)
npm run test:run
```

## Testing Philosophy

These tests focus on:
1. **Infrastructure validation** - Ensuring the test setup works correctly
2. **Core utility functions** - Testing business logic and data transformations
3. **Basic component rendering** - Verifying UI components mount properly
4. **Type safety** - Validating TypeScript interfaces and types
5. **Automatic numbering** - Comprehensive testing of Word document numbering functionality
6. **Extraction numbering behavior** - Coverage for rendered references plus queued restart/override gaps

The tests are designed to be fast, reliable, and focused on the most critical functionality while serving as a foundation for future test expansion.

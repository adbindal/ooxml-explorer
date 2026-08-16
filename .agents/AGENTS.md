# OOXML Explorer: Agent Development Guide & Instructions

Welcome, Agent! You are working on the **OOXML Explorer** repository—a React and TypeScript web application designed to inspect, edit, and compare Office Open XML (OOXML) documents (.docx, .xlsx, .pptx) directly in the browser using JSZip, Monaco Editor, and Google Gemini.

This file serves as your development manual. Read and follow these rules strictly to ensure architectural integrity, stylistic consistency, and clean test compliance.

---

## 1. Architectural Overview & Directories

```
ooxml-explorer/
├── components/          # Reusable UI widgets (FileTree, AIPanel, Logo, etc.)
├── services/            # Core business engines (zipService, geminiService, testService)
├── store/               # State management (Zustand: appStore.ts)
├── views/               # Page-level route views (Landing, Editor, Diff, Validator)
├── utils/               # Formatting, trees, hotkeys, themes, and markdown shims
├── tests/               # Test suites (runs in Vitest and browser shims)
└── types.ts             # Global TS type definitions
```

---

## 2. Core Architecture Rules

### A. State Management (Zustand)
- All shared UI, Editor, and Diff states MUST be managed through the unified Zustand store in [store/appStore.ts](file:///Users/adbindal/code/exp/ooxml-explorer/store/appStore.ts).
- Do NOT create local state for properties that require cross-panel or cross-view synchronization (e.g., active file path, pending changes, or sidebar toggles).
- **Tab State**: When modifying active tabs or files, update `editor.openTabs` and `editor.activePath` in tandem.
- **Resets**: When switching modes back to `'landing'`, always invoke the store's reset logic to purge stale file descriptors and prevent memory leaks.

### B. Zip & Pack Operations (Strict OOXML Compliance)
- All ZIP modifications MUST go through [services/zipService.ts](file:///Users/adbindal/code/exp/ooxml-explorer/services/zipService.ts).
- **Strict Pack Order**: When generating an exported OOXML package, the `mimetype` file **MUST** be placed first in the archive and **MUST NOT** be compressed (`STORE` mode in JSZip). Subsequent XML assets MUST be compressed using `DEFLATE`. Failing this order or compressing the mimetype will result in Microsoft Office rejecting the generated document as corrupted.
- Preserve compression modes of unmodified files when repacking.

### C. Gemini Assistant Integration
- All AI queries, explainers, and structural audits MUST route through [services/geminiService.ts](file:///Users/adbindal/code/exp/ooxml-explorer/services/geminiService.ts).
- Pre-emptively slice file strings to a maximum of 8,000 characters before sending them to the SDK to manage token limits and optimize round-trip latencies.
- Ensure API keys are checked locally via `getApiKey()`. If missing, gracefully throw `API_KEY_MISSING` so the UI can prompt the user to input one in the settings drawer.

### D. Monaco Editor Themes & Word Wrap
- Monaco Editor configurations MUST use the theme definitions declared in [utils/theme.ts](file:///Users/adbindal/code/exp/ooxml-explorer/utils/theme.ts).
- Word wrap (`wordWrap: 'on'`) MUST be enabled for all editor and diff editor models to guarantee readable viewports.

### E. AI Request/Response Validation (Zod)
- Every exported AI service function's request and response shape MUST be a Zod schema (`z.object(...)`), with the TypeScript type derived via `z.infer<typeof Schema>` - never a hand-written `interface` for these shapes. See `EditorFileContextSchema`/`AIAnalysisSchema` in [services/geminiService.ts](file:///Users/adbindal/code/exp/ooxml-explorer/services/geminiService.ts) and `ElementExplanationSchema` in [services/aiService.ts](file:///Users/adbindal/code/exp/ooxml-explorer/services/aiService.ts) as the canonical examples.
- **Validate the request** at the top of the function (`Schema.parse(...)` or `.safeParse(...)`) before doing any work. This catches malformed callers at runtime, not just at compile time.
- **Validate the response** before returning it to the caller. This matters most for the on-device local model, which is not a constrained-decoding API and can return literally anything - Zod validation is what turns "the model returned garbage" into a caught, readable error instead of a silent type mismatch or a crash deeper in the UI. See `promptLocalModelForJson` in `services/geminiService.ts`.
- On a validation failure, throw a short, readable `Error` message - never let a raw `ZodError` (its `.message` is a JSON dump of every issue) reach the UI. Catch the parse call and rethrow with a message the AI panel can display directly.
- Scope: this rule covers the AI service layer's exported request/response contracts - the actual model I/O boundary in `services/aiService.ts` and `services/geminiService.ts` - not every internal helper type in the codebase.

---

## 3. Style Guide & Design Aesthetics

This project uses a highly customized dark-mode first design system built with custom **Tailwind CSS HSL color tokens** mapped in [utils/theme.ts](file:///Users/adbindal/code/exp/ooxml-explorer/utils/theme.ts).
- **Primary Accent**: Use `#4A89DC` (representing standard document editing blue) for primary buttons, active states, and highlights.
- **Custom Selection Colors**: In dark mode, use Indigo/Slate transitions (`text-[#A5B4FC]` / `bg-[#A5B4FC]/10`) to distinguish diff modifications and pending states.
- Do NOT inject hardcoded color classes (like `bg-red-500` or `text-blue-600`) unless they are part of a theme-aware class returned by `useThemeClasses()`.
- Maintain clean animations: use transition-all classes for smooth panel collapses and hover states.

---

## 4. Testing & Verification

We maintain a dual-mode testing strategy:
1. **Vitest CLI**: Run tests programmatically on your terminal during development.
2. **Browser Test Runner**: A custom frontend test harness that executes the exact same test files directly in the browser environment, validating end-to-end browser compatibility.

### Developer Commands
To verify your modifications, run the following commands in the workspace root:

- **Run Linting**:
  ```bash
  npm run lint
  ```
- **Run Unit Tests (Vitest)**:
  ```bash
  npm run test
  ```
- **Run Test Coverage**:
  ```bash
  npm run test:coverage
  ```
- **Build Production Bundle**:
  ```bash
  npm run build
  ```

Always run `npm run lint` and `npm run test` before declaring your implementation complete. If you add new logic, write corresponding test specifications in the `tests/` directory and import them into [services/testService.ts](file:///Users/adbindal/code/exp/ooxml-explorer/services/testService.ts) to register them in the browser runner.

---

## 5. Verification Checklist for Agents

Before completing any task:
1. Run `npm run lint` and fix all static analysis warnings.
2. Run `npm run test` to verify that logic, tree-traversals, zip engines, and stores are fully functional.
3. Build the application with `npm run build` to guarantee there are no TypeScript compile-time errors or missing asset imports.

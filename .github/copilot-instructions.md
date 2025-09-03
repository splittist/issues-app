# Instant Issues List

Instant Issues List is a React + TypeScript web application that extracts interesting paragraphs from `.docx` files and produces a report on them in another `.docx` file. The app uses Vite for building, FluentUI for UI components, and the 'docx' library for Word document processing.

Always reference these instructions first and fallback to search or bash commands only when you encounter unexpected information that does not match the info here.

## Working Effectively

### Bootstrap, Build, and Test
- Install Node.js v20.19.4 or later if not available
- Bootstrap the repository:
  - `npm install` -- takes ~4-15 seconds, installs 245 packages (varies by network)
  - `npm audit fix` -- fixes security vulnerabilities if needed, takes ~10 seconds
- Build the application:
  - `npm run build` -- takes ~7 seconds. TypeScript compilation + Vite bundling
  - Generates production files in `dist/` directory (975KB main bundle)
- Lint the code:
  - `npm run lint` -- takes ~1-2 seconds. Uses ESLint with TypeScript rules
- Run development server:
  - `npm run dev` -- starts Vite dev server on http://localhost:5173/
  - Ready in ~200ms, includes hot module reloading
- Run production preview:
  - `npm run preview` -- serves built files from `dist/` on http://localhost:4173/

### Validation
- ALWAYS manually validate changes by running the development server and testing file upload/processing
- **Critical validation scenario**: Upload a `.docx` file with highlighted text, comments, or redlining to test extraction functionality
- **UI validation**: Test checkbox interactions for extraction criteria (redlining, highlighting, square brackets, etc.)
- **File processing validation**: Verify that the "Save file" button generates and downloads a report file
- **Output validation**: Check that extracted content appears correctly in the generated .docx report
- ALWAYS run `npm run lint` before committing changes or CI will fail
- No automated tests exist in this project - manual testing is required for all functionality

## Application Architecture

### Key Components
- **App.tsx** (21 lines): Main application shell with header and footer
- **WordHandler.tsx** (125 lines): Core component handling file upload, criteria selection, and processing
- **FileItem.tsx** (77 lines): Individual file item with drag-and-drop reordering
- **wordUtils.ts** (1154 lines): Heavy lifting for .docx parsing and extraction logic
- **types.ts** (61 lines): TypeScript interfaces and types for the application

### Core Functionality
The application allows users to:
1. Upload one or more `.docx` files via drag-and-drop or file selection
2. Configure extraction criteria:
   - Redlining (track changes) - enabled by default
   - Highlighted text
   - Square brackets content
   - Comments
   - Footnotes
   - Endnotes
3. Specify output filename (defaults to `report_YYYY-MM-DD.docx`)
4. Generate and download a report containing extracted paragraphs

### Technology Stack
- **React 18.3.1**: UI framework
- **TypeScript 5.6.2**: Type-safe JavaScript
- **Vite 6.3.5**: Build tool and dev server
- **FluentUI**: Microsoft's React component library
- **docx 9.1.0**: Word document creation and parsing
- **react-dropzone 14.3.5**: File drag-and-drop
- **react-dnd 16.0.1**: Drag-and-drop for file reordering
- **file-saver 2.0.5**: Client-side file downloading

## Common Tasks

### Repository Structure
```
/home/runner/work/issues-app/issues-app/
├── .github/                    # GitHub configuration
├── src/                        # Source code
│   ├── App.tsx                 # Main application
│   ├── WordHandler.tsx         # File processing component
│   ├── FileItem.tsx            # File list item
│   ├── wordUtils.ts            # Document parsing logic
│   ├── types.ts                # TypeScript definitions
│   └── utils.ts                # Utility functions
├── dist/                       # Build output (generated)
├── node_modules/               # Dependencies (generated)
├── package.json                # Project configuration
├── vite.config.ts              # Vite configuration
├── tsconfig.json               # TypeScript configuration
├── eslint.config.js            # ESLint configuration
└── index.html                  # HTML template
```

### Build Configuration Files
- **vite.config.ts**: Minimal Vite config with React plugin
- **tsconfig.json**: References app and node TypeScript configs
- **eslint.config.js**: ESLint with TypeScript, React hooks, and React refresh rules
- **package.json**: ES modules enabled (`"type": "module"`)

### Dependency Management
- 21 production dependencies (React, FluentUI, docx library, etc.)
- 14 development dependencies (TypeScript, ESLint, Vite, etc.)
- Audit vulnerabilities automatically fixable with `npm audit fix`
- No security issues after running audit fix

### Development Workflow
1. Start with `npm install` to install dependencies
2. Run `npm run dev` to start development server
3. Make changes to source files (automatic hot reload)
4. Run `npm run lint` to check code style
5. Run `npm run build` to verify production build works
6. Test functionality manually by uploading .docx files

### File Processing Details
The application processes Word documents by:
1. Parsing .docx files using the docx library
2. Extracting paragraphs based on selected criteria
3. Building a new document with extracted content
4. Providing download of the generated report

Key extraction criteria:
- **Redlining**: Track changes/revisions in documents
- **Highlighting**: Text with background highlighting
- **Square brackets**: Content within [brackets]
- **Comments**: Word document comments
- **Footnotes/Endnotes**: Referenced notes

### Common Issues and Solutions
- **Build warnings**: Large bundle size (975KB) is expected due to Word processing libraries
- **Module errors**: Ensure Node.js v20+ is used (project uses ES modules)
- **Audit vulnerabilities**: Run `npm audit fix` after fresh install
- **Type errors**: Check that all TypeScript interfaces in `types.ts` are properly imported

## Performance Notes
- **npm install**: ~15 seconds for 245 packages
- **npm run build**: ~7 seconds (TypeScript + Vite bundling)
- **npm run lint**: ~1-2 seconds
- **npm run dev**: ~200ms startup time
- **File processing**: Depends on document size and complexity

Always test the complete file upload → processing → download workflow when making changes to ensure the core functionality remains intact.
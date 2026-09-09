# React migration note

You asked to migrate the add-in UI to React. I added a lightweight React skeleton under `office-addin/src-react` and a `package.react.json` with scripts for vite-based development. This is a starting point; I'll replace the existing plain JS taskpane with the React build and hook Office.js APIs into the React components in the next steps.

Next steps:
- Wire Office.js selection events and metadata API into the React components.
- Replace the manifest SourceLocation during dev to point to the built React app.
- Update CI/build scripts and finalize the one-time delivery bundle.


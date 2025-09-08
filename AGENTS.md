# Agent Instructions

## Manifest Validation

After making any changes to `manifest.xml`, you **must** validate it to ensure it conforms to the schema.

Run the following command to validate the manifest:

```bash
npm run validate
```

If the validation fails, you must correct the `manifest.xml` file before submitting your changes.

## Code Linting and Formatting

After making any changes to any typescript, html, or css files, you **must** format and lint it to ensure it still passes

Run the following command to do so:

```bash
npm run prettier
npm run lint
```

If the linting fails, you must correct the files before submitting your changes.

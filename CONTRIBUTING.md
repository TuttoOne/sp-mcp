# Contributing to SharePoint Bridge for Claude

Thank you for your interest in contributing. This project is maintained by [Tutto.one](https://tutto.one) and we welcome contributions from the community.

## How to Contribute

### Reporting Bugs

1. Check existing [Issues](https://github.com/TuttoOne/sp-mcp/issues) to avoid duplicates
2. Open a new issue with:
   - Clear title describing the problem
   - Steps to reproduce
   - Expected vs. actual behaviour
   - Your environment (Node.js version, OS, SharePoint tier)
   - Relevant error messages or logs (redact any credentials)

### Suggesting Features

1. Open an issue with the `enhancement` label
2. Describe the use case — what problem does this solve?
3. If possible, describe how it might work from the user's perspective

### Submitting Code

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/your-feature-name`
3. Make your changes
4. Build and verify: `npm run build`
5. Test against a real SharePoint environment (or document that you couldn't)
6. Submit a Pull Request with:
   - Clear description of what changed and why
   - Any new environment variables or configuration needed
   - Any new Graph API permissions required

### Industry Templates

We especially welcome industry template contributions — if you've built a SharePoint architecture for a specific sector (legal, property management, healthcare, finance, recruitment, etc.) and want to share the schema, we'd love to include it.

A template contribution includes:
- List designs (names, descriptions)
- Column definitions (names, types, choices)
- Relationship map (which lists link to which via lookups)
- Optionally: sample data and Claude prompt templates

Open an issue describing the template and we'll collaborate on getting it into the project.

## Code Style

- TypeScript throughout
- Use `const` and `let`, never `var`
- Async/await over raw Promises
- Meaningful variable names
- Comments for anything non-obvious
- Error messages should be helpful to the end user (who may not be a developer)
- All config from environment variables, never hardcoded

## Project Structure

```
sp-mcp/
├── src/
│   ├── index.ts          # MCP server entry point
│   ├── tools/            # Tool implementations (SharePoint, PA, Docs)
│   └── services/         # Graph API client, auth, helpers
├── dist/                 # Compiled output (git-ignored)
├── .env.example          # Environment variable template
├── ecosystem.config.cjs  # PM2 configuration
├── package.json
├── tsconfig.json
├── README.md
├── LICENSE
└── CONTRIBUTING.md
```

## Security

If you discover a security vulnerability, **do not open a public issue**. Email security@tutto.one with details and we'll address it promptly.

## Questions?

Open a Discussion on GitHub or email daniel@tutto.one.

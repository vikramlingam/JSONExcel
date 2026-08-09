# Security

## Reporting a vulnerability

Use the repository's private vulnerability-reporting feature when available. If it is unavailable,
contact a repository maintainer privately. Do not open a public issue containing exploit details,
credentials, or sensitive sample data.

Include the affected version, reproduction steps, impact, and any suggested mitigation. Maintainers
will acknowledge the report, investigate it, and coordinate disclosure and remediation.

## Security defaults

`jsonexcel` writes formula-like strings as text by default to reduce formula-injection risk. Do
not enable formula writing for untrusted input. URL and email hyperlinks are opt-in and use
conservative patterns. See [docs/security.md](docs/security.md) for operational guidance.

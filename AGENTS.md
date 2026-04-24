# AGENTS.md

## General rules

- Prefer small, focused changes.
- Do not change public APIs unless the issue explicitly requires it.
- Add or update tests for every behavior change.
- Do not weaken encryption settings or compatibility behavior.
- Keep test data deterministic.

## Testing requirements

Before completing any issue, run:

```bash
dotnet test
```

## Compatibility principle

This project values compatibility over cleverness.

When changing encryption/decryption behavior, prefer explicit tests against known fixtures or third-party implementations such as Apache POI.

Do not consider an implementation correct only because ExcelEncryptor can decrypt files encrypted by itself.



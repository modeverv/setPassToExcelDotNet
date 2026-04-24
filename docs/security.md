# Security Notes

This library implements Excel-compatible MS-OFFCRYPTO Agile Encryption for OOXML workbooks (`.xlsx`, `.xlsm`).

## What this library protects

- Protects workbook files at rest using password-based encryption.
- Uses Office Agile Encryption parameters compatible with Microsoft Excel and Apache POI.

## Recommended setting

- Recommended for new encrypted files: `Aes256` + `Sha512`.

## Compatibility settings

- Legacy settings such as `Sha1` and `Md5` are provided for compatibility scenarios.
- Legacy settings are not recommended for new files.

## Security boundaries and limitations

- Password strength is the responsibility of the user.
- After a workbook is opened by Excel or another application, in-memory contents are not protected by this library.
- This library is not a DRM solution.
- This library does not prevent re-save, copy, or redistribution after the workbook has been opened.


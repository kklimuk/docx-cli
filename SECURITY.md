# Security Policy

## Reporting a vulnerability

If you discover a security vulnerability in docx-cli, please report it privately:

- Open a [GitHub security advisory](https://github.com/kklimuk/docx-cli/security/advisories/new), or
- Email **kklimuk@gmail.com**.

Please don't open a public issue for security-sensitive reports. We aim to acknowledge within 72 hours and to ship a fix or mitigation promptly.

## Supported versions

Security fixes target the **latest published release**. Older versions are not maintained — upgrade with:

```sh
curl -fsSL https://raw.githubusercontent.com/kklimuk/docx-cli/main/install.sh | sh
# or
bun add -g bun-docx
```

## Scope and data handling

docx-cli runs entirely **locally** against `.docx` files on disk and transmits no document content anywhere. The only network activity is:

- **`docx render`** — shells out to a locally installed Word (macOS/Windows) or LibreOffice to produce a PDF; no data leaves the machine.
- **`skills/docx-cli/scripts/bootstrap.sh`** and **`install.sh`** — fetch the prebuilt `docx` binary from this repo's GitHub Releases over HTTPS (binary download only).

Mutating commands overwrite the target file in place (git is the history); there is no telemetry and no external API.

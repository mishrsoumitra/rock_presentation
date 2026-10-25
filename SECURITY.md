# Security

Presentation files are untrusted input. The parsers, the editor and the server are built for that: limits are checked
before anything is inflated, every upload is parsed in a worker thread with a heap cap and a deadline, slides are
drawn in a Chromium page that can reach nothing outside the server, and the editor runs under a content security
policy. README.md describes these defences, and the tests exercise them with fuzzing, ZIP bombs and malformed files.

## Reporting a vulnerability

Please report it privately through GitHub: the repository's **Security** tab, **Report a vulnerability**. Include a
file or the steps that show the problem if you can. Please do not open a public issue for it.

You should hear back within a week. Fixes go to `main` and into the next container image.

## Supported versions

Only the latest `main` and the images built from it (`ghcr.io/mishrsoumitra/rock_presentation`) get fixes.

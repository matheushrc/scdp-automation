# Chrome profile selection and project clone design

## Goal

Let an operator choose an existing Google Chrome profile from an arrow-key terminal menu, copy only that profile into the repository-local `.scdp-browser` directory, and run the SCDP command using the copied profile. Support the Fedora/Linux environment already used by the project and the Windows environment used by the operator's manager, without adding a Python dependency.

## Proposed scope

- Discover Chrome Stable profiles for the current operating-system user. On Windows, use `%LOCALAPPDATA%\Google\Chrome\User Data`; on Linux, use Chrome's configuration directory, honoring `CHROME_CONFIG_HOME` and `XDG_CONFIG_HOME` when set. Chromium documents these paths and the user-data-directory/profile relationship in its [User Data Directory guide](https://chromium.googlesource.com/chromium/src/%2B/refs/heads/main/docs/user_data_dir.md).
- Read display names and profile-directory keys from the root `Local State` file and include only recognized profile directories that exist (`Default` and `Profile N`). If metadata for a directory is absent, use its directory name as the label. Read the profile `name` for the menu but do not read or log account-name/email metadata; if the display name itself looks like an email address, show the directory name instead.
- Provide an interactive menu with Up/Down and Enter, and Escape to cancel. Use Python's standard library (`msvcrt` on Windows and `termios`/`tty` on POSIX), with no new dependency. Fail with a clear terminal requirement when standard input or output is not interactive.
- If `.scdp-browser` is missing or incomplete, show the menu before running extraction, `--login`, or `--abrir-navegador`. After copying, continue with the requested command. Reuse a complete project clone on ordinary runs so authentication and state accumulated in that clone persist. Add `--selecionar-perfil-chrome` to explicitly choose and copy a different source profile before running the requested command. This first-run/reselect cadence is a design assumption for review.
- Support launching the visible installed Chrome on Windows as well as Linux. Continue to bind the remote-debugging endpoint to loopback, keep extensions disabled, and use the selected cloned profile.

## Profile-copy behavior

The selected Chrome profile is copied into the repository root at `.scdp-browser`; the source profile is never modified. The clone contains only the selected profile directory, root `Local State`, and a `.scdp-profile-directory` marker recording the selected directory name. Keep the selected directory name rather than renaming profile data, and make the launcher use that directory. Filter `Local State` profile metadata to the selected profile and update its last-used references so the clone does not advertise profiles whose data was not copied. Preserve all unrelated `Local State` values, including Chrome's encryption state. Existing clones without the marker continue to use `Default`.

Before copying, require Chrome to be closed and check for running Chrome processes using platform-appropriate facilities. If Chrome is open, report that the operator must close its windows and retry; do not copy a profile while Chrome may be writing it. Copy into a staging directory beside `.scdp-browser`, validate the staged profile and metadata, then replace the destination. If a destination already exists, first save it as `.scdp-browser.backup-<timestamp>`. A failure before replacement leaves the current clone intact; a failure during replacement restores it from the staging/backup state.

If `Local State` is missing or invalid, no usable profile directories exist, or staging fails, abort with a concise error and leave the source and existing clone untouched. Do not print profile metadata, account names, or file contents in errors or logs.

Do not copy extension payload or extension storage directories into the clone. The existing Chrome launch flags continue to disable extensions. The source's other profile directories are not copied. `.gitignore` already excludes `.scdp-browser*`; keep the clone, staging data, and backups local and out of version control.

## CLI and browser-launch integration

Keep the current default extraction command and its `--limite` handling. Profile preparation must run before the selected action so all three paths—extraction, `--login`, and `--abrir-navegador`—use the same clone. Existing clones without `.scdp-profile-directory` continue to use `Default`.

Resolve Chrome Stable from `PATH` and standard install locations. On Windows, check the current user's Local AppData installation and machine-level Program Files locations for `chrome.exe`. On Linux, retain the current `google-chrome` lookup and support the common `google-chrome-stable` executable name. Pass arguments as separate subprocess arguments so paths containing spaces work. Apply POSIX-only process options only on POSIX systems; do not pass `start_new_session` on Windows.

The profile copy is local to the account running the script. It does not enumerate Chrome data belonging to other Windows accounts, and it does not move a profile between computers. Chrome uses application- and machine-bound protection for sensitive Windows profile data; Chromium also checks whether Chrome is using its default user-data directory when deciding whether App-Bound encryption is supported ([implementation](https://chromium.googlesource.com/chromium/src/%2B/main/chrome/browser/os_crypt/app_bound_encryption_win.cc)). The implementation will not decrypt or bypass that protection, and it cannot guarantee that a copied session remains usable. The copied profile must be checked on the target Windows machine; if gov.br requests login or CAPTCHA, the operator completes it in the visible browser. Google's [Chrome cookie-encryption explanation](https://security.googleblog.com/2024/07/improving-security-of-chrome-cookies-on.html) describes the Windows application-bound protection.

## Validation

- Unit-test profile discovery with temporary `Local State` fixtures: multiple profiles, missing names, missing directories, malformed JSON, unsupported operating systems, and both Windows and Linux path resolution.
- Unit-test arrow-menu selection by injecting key reads and output streams, including wraparound, Enter, Escape, and non-interactive terminal errors.
- Unit-test copying: only the selected source profile is present; extension directories are absent; the clone metadata points to the selected directory; existing clones are backed up before replacement; interrupted/failed staging does not destroy the current clone; and no source files are changed.
- Unit-test CLI behavior: first-run setup, reuse of a complete clone, explicit reselection, and profile preparation before extraction, login, or browser-only startup.
- Unit-test browser executable resolution and launch arguments for Linux and Windows without launching a real browser in unit tests.
- Run `uv run python -m unittest discover -v`, `uv run ruff check .`, `uv run ruff format --check .`, and `uv run ty check` from the repository root.
- On Windows, manually test profile discovery, selection, copying, visible Chrome startup, and whether the copied profile reaches the authenticated SCDP page. Do not automate or bypass gov.br challenges. The current Fedora environment can test Windows path and launch logic through mocks but cannot validate Windows Chrome encryption or its GUI.

## Out of scope

- Listing profile data from other operating-system accounts, Chrome Beta/Dev/Canary, Chromium, or Chrome for Testing.
- Prompting and copying on every ordinary run, or automatically synchronizing changes between the source profile and project clone.
- Reading saved passwords, exporting cookies, bypassing Chrome's encryption, automating CAPTCHA, or changing gov.br authentication behavior.
- Supporting profile copying across machines or Windows accounts.
- Changing extraction concurrency, tab behavior, output format, or SCDP report handling.

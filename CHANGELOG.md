# Changelog

All notable changes to this project are documented here. The format is based on
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project follows
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Fixed
- Markdown to Teams HTML: a blank line between paragraphs now renders as a visible gap. Teams' `RichText/Html` chat renderer collapses the margin between adjacent `<p>` elements, so paragraph-to-paragraph breaks are now emitted as `<br><br>` inside a single `<p>`. Genuine block elements (lists, tables, headings, code, blockquotes) keep their own boundaries, and a heading sitting directly above its content stays tight.
- Login / token extraction against MSAL Browser v4 encrypted `localStorage` cache: decrypt entries using the `msal.cache.encryption` session cookie (AES-GCM + HKDF), matching `@azure/msal-browser`. Without this, Substrate search and HTTP token refresh could fail after a successful login while messaging cookies still worked.

## [0.29.2] - 2026-07-01

### Fixed
- Markdown to Teams HTML: a line consisting solely of a bold label (e.g. `**Target**`) is now rendered as its own paragraph instead of being joined to the following line with a `<br>`, so a heading sits directly above its content without needing a blank-line workaround. Handles indented labels and trailing hard-break markers correctly.
- `teams_find_meeting_times` now returns unambiguous UTC timestamps. It previously returned times in the mailbox's local timezone with no offset marker, which could be misread as UTC and lead to meetings suggested or scheduled at the wrong time for anyone not on a UTC mailbox.
- HTTP 403 responses are now classified as a permissions error rather than an authentication error. Previously a genuine "you don't have permission" failure (e.g. deleting a channel you don't own, deleting someone else's message without moderator rights) triggered an unnecessary automatic re-login attempt followed by a misleading "call teams_login" suggestion.
- HTTP rate limiting is now tracked per API host instead of globally, so a 429 from one API (e.g. the calendar API) no longer blocks unrelated calls to other APIs (e.g. search or messaging) for the duration of the `Retry-After` window. The client also now honours the server's `Retry-After` value for its own retries (previously it fell back to generic exponential backoff regardless), and the recorded rate-limit window is correctly capped to match the delay actually used.
- Token refresh no longer discards already-successfully-refreshed tokens when a later scope fails with an auth error (e.g. missing consent for one specific API). Previously this silently threw away progress made on the other scopes and forced a slower browser-based re-authentication fallback even when most tokens had already refreshed successfully.
- Fixed several smaller robustness issues found during a codebase review: a missing barrel export for task and channel management tools, a browser process left running after a failed authentication attempt, an HTML entity decode-order bug that could over-decode double-encoded content, an unguarded `decodeURIComponent` call that could throw on a malformed session cookie, and unbounded `maxResults` search input that could silently drop results instead of erroring on invalid input.

### Changed
- CI pipeline: lint, typecheck and build now run once in a dedicated job instead of once per Node version; the test matrix now covers Node 22, 24 and 26 (dropped Node 20, which is past active support) and depends on the lint/build job passing first; test coverage is now enforced in CI; `actions/checkout` and `actions/setup-node` updated to their latest major versions.
- Removed the automated AI PR-reviewer workflow, which depended on GitHub App credentials that were never configured for this fork and caused every pull request to show a failing check.

## [0.29.1] - 2026-06-23

### Added
- `teams_get_message` now returns `rawHtml` (original HTML content before stripping) and `rawFileObjects` (the parsed file chiclet array from `properties.files`), making it possible to inspect or re-post a message's exact content and attachments.

### Changed
- `teams_forward_message` now carries file attachments from the source message into the target conversation as native chiclets, with no download or re-upload. The original file objects are passed directly, preserving the original filename and SharePoint URL. The hardcoded "Forwarded message:" label has been removed — only the quoted block (sender name + content) is included by default. The optional `comment` parameter still prepends a note above the quote.

## [0.29.0] - 2026-06-23

### Added
- `teams_download_shared_file`: download a file shared in a Teams conversation by its SharePoint/OneDrive URL (the `webUrl` from `teams_get_shared_files`). Works for files from other users' drives that you have been granted access to, using the Microsoft Graph Shares API. Use `teams_download_file` for files in your own OneDrive.

## [0.28.0] - 2026-06-18

### Added
- `teams_send_files`: send several local files into a conversation as native attachments on a **single message** (one post, multiple file chiclets), with an optional caption. Use it instead of calling `teams_send_file` once per file when the files should be grouped under one message. Files are uploaded to the right place per conversation type, the same as `teams_send_file`.

## [0.27.0] - 2026-06-16

### Added
- `teams_get_recording`: get a meeting's recording(s) from the meeting `threadId` (from `teams_get_meetings`). Returns metadata (title, file name, start/end time, duration) plus a `playbackUrl` (opens in the SharePoint/Stream web player) and a `downloadUrl` (direct `.mp4`) for each recording, rather than the video bytes. Uses the same Substrate `WorkingSetFiles` lookup as `teams_get_transcript`, so no extra authentication is needed.

## [0.26.1] - 2026-06-15

### Fixed
- Markdown to Teams HTML: a list, blockquote or table now renders correctly when it immediately follows a text line within the same block (e.g. a `Label:` line directly above `- bullets`), without needing a separating blank line. The converter previously required every line in a block to share one type, so a label plus a list fell back to a single `<br>`-joined paragraph and the bullets did not render. Blocks are now parsed by grouping consecutive line runs, letting block elements interrupt a paragraph as standard markdown allows.

## [0.26.0] - 2026-06-09

### Added
- `teams_send_file` now sends a **native file attachment** (a real file chiclet that also shows in the conversation's Files tab) instead of posting a share link. The file is uploaded to the right place per conversation type, matching the Teams client: a channel's own SharePoint files folder for channels, or OneDrive "Microsoft Teams Chat Files" (shared via an organisation link) for 1:1, group, meeting and self chats.
- `teams_create_channel` and `teams_delete_channel` to create and delete team channels (standard or private) via the middle-tier API.

## [0.25.1] - 2026-06-04

### Changed
- Documentation only: added npm badges (version, downloads, node, license) and pointed all clone, install and repository references at `@shayanline/msteams-mcp`.

## [0.25.0] - 2026-06-04

### Added
- Messaging: `importance` (high/urgent) and `subject` on `teams_send_message`, plus `teams_forward_message`.
- Browse: `teams_list_chats`, `teams_list_teams`, `teams_mark_unread`, `teams_get_chat_members`.
- Chat management: `teams_rename_chat`, `teams_add_member`, `teams_remove_member`, `teams_leave_chat`, `teams_pin_message`, `teams_unpin_message`, `teams_mute_chat`, `teams_unmute_chat`.
- Scheduling: `teams_find_meeting_times`.
- Files: `teams_list_files`, `teams_upload_file`, `teams_download_file`, `teams_send_file`.
- Microsoft To Do: `teams_list_task_lists`, `teams_list_tasks`, `teams_create_task`, `teams_complete_task`.

### Changed
- Project is now maintained and released independently as `@shayanline/msteams-mcp`. `main` is the source of truth.

## [0.24.1] and earlier

See the [GitHub releases](https://github.com/shayanline/msteams-mcp/releases) for the history before independent maintenance began.

[Unreleased]: https://github.com/shayanline/msteams-mcp/compare/v0.29.2...HEAD
[0.29.2]: https://github.com/shayanline/msteams-mcp/compare/v0.29.1...v0.29.2
[0.29.1]: https://github.com/shayanline/msteams-mcp/compare/v0.29.0...v0.29.1
[0.29.0]: https://github.com/shayanline/msteams-mcp/compare/v0.28.0...v0.29.0
[0.28.0]: https://github.com/shayanline/msteams-mcp/compare/v0.27.0...v0.28.0
[0.27.0]: https://github.com/shayanline/msteams-mcp/compare/v0.26.1...v0.27.0
[0.26.1]: https://github.com/shayanline/msteams-mcp/compare/v0.26.0...v0.26.1
[0.26.0]: https://github.com/shayanline/msteams-mcp/compare/v0.25.1...v0.26.0
[0.25.1]: https://github.com/shayanline/msteams-mcp/compare/v0.25.0...v0.25.1
[0.25.0]: https://github.com/shayanline/msteams-mcp/compare/v0.24.1...v0.25.0
